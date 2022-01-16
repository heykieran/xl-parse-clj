(ns graph
  (:require
   [clojure.string :as str]
   [ubergraph.core :as uber]
   [ubergraph.alg :as alg]
   [dk.ative.docjure.spreadsheet :as dk]
   [clojure.math.numeric-tower :as math]
   [xlparse :as parse]
   [excel :as excel]
   [ast-processing :as ast]
   [shunting :as sh]
   [box :as box]
   [clojure.walk :as walk])
  (:import
   [org.apache.poi.ss SpreadsheetVersion]
   [org.apache.poi.ss.util AreaReference]
   [org.apache.poi.ss.usermodel CellType DateUtil Name SheetVisibility]
   [org.apache.poi.xssf.usermodel XSSFName XSSFEvaluationWorkbook]
   [org.apache.poi.ss.formula FormulaParser FormulaType]
   [org.apache.poi.ss.formula.ptg Ptg]
   [box Box]))

(declare ^:dynamic *context*)

(defn range-metadata [cell-range-str]
  (let [aref (AreaReference. cell-range-str SpreadsheetVersion/EXCEL2007)
        f-cell (.getFirstCell aref)
        l-cell (.getLastCell aref)
        f-row (.getRow f-cell)
        f-col (.getCol f-cell)
        l-row (.getRow l-cell)
        l-col (.getCol l-cell)
        [cell-sheet-name cell-row-name cell-col-name] (.getCellRefParts f-cell)]
    {:areas [{:single? (.isSingleCell aref)
              :column? (.isWholeColumnReference aref)
              :sheet-name cell-sheet-name
              :tl-name (str cell-col-name cell-row-name)
              :tl-coord [f-row f-col]
              :cols (inc (apply - (sort > [l-col f-col])))
              :rows (inc (apply - (sort > [l-row f-row])))}]}))

(defn get-named-ranges-from-wb-map
  "For each sheet access its named ranges and consolidate
   into a single sequence."
  [wb-map]
  (mapcat
   (fn [[sheet-name {:keys [named-ranges] :as s-map}]]
     named-ranges)
   wb-map))

(defn get-named-range-description-map [named-ranges cell-name]
  (tap> {:loc get-named-range-description-map
         :named-ranges named-ranges
         :cell-name cell-name})
  (when (and named-ranges (-> cell-name
                              (AreaReference. SpreadsheetVersion/EXCEL2007)
                              (.isSingleCell)))
    (some (fn [{:keys [name sheet] :as named-range}]
            (when (= cell-name (str sheet "!" name))
              named-range))
          named-ranges)))

(defn expand-cell-range
  ([cell-range-str]
   (expand-cell-range cell-range-str nil))
  ([cell-range-str wb-map]
   (letfn [(cell-info [ref-cell]
             (let [[sheet-name row-name col-name]
                   (.getCellRefParts ref-cell)]
               {:sheet sheet-name
                :label (str col-name row-name)
                :type :general}))]
     (if-let [named-range (-> wb-map
                              (get-named-ranges-from-wb-map)
                              (get-named-range-description-map cell-range-str))]
       (:references named-range)
       (let [aref (AreaReference. cell-range-str SpreadsheetVersion/EXCEL2007)]
         (-> (mapv cell-info (.getAllReferencedCells aref))
             (with-meta (graph/range-metadata cell-range-str))))))))

(defn explain-sheet-cells [sheet-name cells]
  (reduce
   (fn [accum cell]
     (if (nil? cell)
       accum
       (let [value-type (excel/get-cell-type cell)
             cell-style (.getCellStyle cell)
             cell-address (.getAddress cell)]
         (conj accum
               (let [cell-value (case value-type
                                  :numeric (.getNumericCellValue cell)
                                  :string (.getStringCellValue cell)
                                  :boolean (.getBooleanCellValue cell)
                                  :empty nil
                                  :error "#ERROR"
                                  "")
                     look-like-date? (and (= :numeric value-type)
                                          (DateUtil/isCellDateFormatted cell))]
                 (cond->
                  {:type value-type
                   :sheet sheet-name
                   :formula (when (= CellType/FORMULA (.getCellType cell)) (.getCellFormula cell))
                   :format (.getDataFormatString cell-style)
                   :label (.formatAsString cell-address)
                   :row (.getRow cell-address)
                   :column (.getRow cell-address)
                   :value cell-value}
                   look-like-date?
                   (merge
                    {:excel-date-value (.getDateCellValue cell)
                     :calc-date-value (excel/excel-serial-date->local-date-time cell-value)})))))))
   []
   cells))

(defn add-references [sheet-name named-ranges cells]
  (mapv (fn [{:keys [formula address value] :as cell}]
          (if formula
            (assoc cell
                   :references
                   (->> (parse/parse-to-tokens (str "=" formula))
                        (keep (fn [{:keys [value type sub-type] :as token}]
                                (when (and (= sub-type :Range)
                                           (= type :Operand))
                                  (expand-cell-range (str sheet-name "!" value) named-ranges))))
                        (mapcat (fn [expanded-ranges]
                                  expanded-ranges))
                        (distinct)
                        (into [])))
            cell))
        cells))

(defn explain-named-ranges-in-workbook
  "Naive attempt to get information about all the named ranges in the workbook.
   Currently can't handle hidden named ranges which causes problems.
   Returns a map with keys :named-ranges and :warnings, each containing
   a vector of information about the named ranges in the WB.
   :warnings will always be nil."
  [wb-as-resource]
  (->>
   wb-as-resource
   (.getAllNames)
   (map (fn [n]
          {:name (.getNameName n)
           :sheet (.getSheetName n)
           :references (expand-cell-range (.getRefersToFormula n))}))
   ((fn [n-ranges]
      {:named-ranges n-ranges
       :warnings nil}))))

(defn explain-named-ranges-in-workbook-ext
  "Get information about all the named ranges in the workbook.
   Returns a map with keys :named-ranges and :warnings, each containing
   a vector of information about the named ranges in the WB.
   :warnings is used to contain information about named ranges that we
   don't know how to handle yet."
  [wb-as-resource]
  (let
   [evaluation-wb (->> wb-as-resource (XSSFEvaluationWorkbook/create))
    ;; use reflection to get a reference to the getCTCName method, which
    ;; is private
    name-hidden-method (-> XSSFName
                           (.getDeclaredMethod
                            "getCTName"
                            (into-array Class nil))
                           (doto (.setAccessible true)))]
    (->>
     wb-as-resource
     (.getAllNames)
     (reduce (fn [accum ^Name n]
               ;; ignore deleted names and hidden names
               (if (and (not (.isDeleted n))
                        (not (-> name-hidden-method
                                 (.invoke n (into-array Object nil))
                                 (.getHidden))))
                 ;; parse the formula, future work so we can
                 ;; use named ranges pointing to values rather to other
                 ;; cells. Currently, any weirdness will end up in the 
                 ;; :warning value of the returned map
                 (let [ptgs (FormulaParser/parse
                             (.getRefersToFormula n)
                             evaluation-wb
                             FormulaType/CELL
                             (.getSheetIndex n))
                       sheet-visibility
                       (if (not= -1 (.getSheetIndex n))
                         (-> wb-as-resource
                             (.getSheetVisibility
                              (.getSheetIndex n))
                             ((fn [^SheetVisibility vis]
                                (cond
                                  (= vis SheetVisibility/HIDDEN) :hidden
                                  (= vis SheetVisibility/VERY_HIDDEN) :very-hidden
                                  (= vis SheetVisibility/VISIBLE) :visible))))
                         :not-applicable)
                       ;; see if we can convert the formula to a contiguous area, returns
                       ;; nil if we can't
                       contiguous-area-ref (try (AreaReference/generateContiguous
                                                 SpreadsheetVersion/EXCEL2007
                                                 (.getRefersToFormula n))
                                                (catch Exception e
                                                  nil))
                       ;; get a list of refed sheets, check later that only
                       ;; one is returned
                       refed-sheets (when contiguous-area-ref
                                      (mapcat
                                       (fn [a]
                                         (reduce (fn [accum c]
                                                   (let [[s _ _] (.getCellRefParts c)]
                                                     (conj accum s)))
                                                 #{}
                                                 (.getAllReferencedCells a)))
                                       contiguous-area-ref))
                       refed-sheet-names (mapv (fn [s]
                                                 (try (-> wb-as-resource
                                                          (.getSheet s)
                                                          (.getSheetName))
                                                      (catch Exception e
                                                        "##INVALID SHEET##")))
                                               refed-sheets)]
                   ;; if everything looks normal, add what we know about the 
                   ;; named range to the :named-ranges vector
                   (if (and (= 1 (count contiguous-area-ref))
                            (= 1 (count refed-sheets))
                            (= 1 (count refed-sheet-names))
                            (= 1 (count ptgs)))
                     (update accum :named-ranges
                             (fnil conj [])
                             {:name (.getNameName n)
                              :contiguous-ref (-> contiguous-area-ref (first) (.formatAsString))
                              :visibility sheet-visibility
                              :formula (.getRefersToFormula n)
                              :function (.getFunction n)
                              :index (.getSheetIndex n)
                              :refers-to-sheets (first refed-sheets)
                              :refers-to-sheet-names (first refed-sheet-names)
                              :parsed-formula (->>
                                               ptgs
                                               (mapv identity)
                                               (first))
                              :refers-to-deleted-cells (Ptg/doesFormulaReferToDeletedCell ptgs)
                              :sheet-name-for-index (try (.getSheetName n) (catch Exception e :no-name))
                              :references (graph/expand-cell-range (.getRefersToFormula n))})
                     ;; something deosn't look right so add what we know to the
                     ;; :warnings vector
                     ;; TODO: If we find a named range that refers to a value or
                     ;; a formula, it will end up here. I need to make sure we can
                     ;; use those too
                     (update accum :warnings
                             (fnil conj [])
                             {:name (.getNameName n)
                              :contiguous-ref contiguous-area-ref
                              :visibility sheet-visibility
                              :formula (.getRefersToFormula n)
                              :function (.getFunction n)
                              :index (.getSheetIndex n)
                              :refers-to-sheets refed-sheets
                              :refers-to-sheet-names refed-sheet-names
                              :parsed-formula (->>
                                               ptgs
                                               (mapv identity))
                              :refers-to-deleted-cells (Ptg/doesFormulaReferToDeletedCell ptgs)
                              :sheet-name-for-index (try (.getSheetName n) (catch Exception e :no-name))
                              :references nil})))
                 accum))
             {}))))

(defn explain-cells-in-sheet [wb-as-resource sheet-name]
  (let [named-ranges (-> wb-as-resource
                         (explain-named-ranges-in-workbook-ext)
                         (:named-ranges))]
    (->> wb-as-resource
         (dk/select-sheet sheet-name)
         dk/cell-seq
         (explain-sheet-cells sheet-name)
         (add-references sheet-name {sheet-name {:named-ranges named-ranges}})
         ((fn [r]
            {:named-ranges named-ranges
             :cells r})))))

(defn explain-workbook
  ([wb-name & [sheet-name]]
   (let [wb-as-resource (dk/load-workbook-from-resource wb-name)
         sheet-names (->> wb-as-resource
                          (dk/sheet-seq)
                          (keep (fn [xl-sheet]
                                  (let [s-name (.getSheetName xl-sheet)]
                                    (when (or (nil? sheet-name)
                                              (= s-name sheet-name))
                                      s-name)))))]
     (->> sheet-names
          (reduce
           (fn [accum sheet-name]
             (assoc accum sheet-name
                    (explain-cells-in-sheet wb-as-resource sheet-name)))
           {})))))

(defn get-cell-dependencies-for-sheet
  "For an individual excel sheet, expressed as a map, say returned by explain-workbook,
   which is a map relating the sheet name to the cells in the sheet, this function returns 
   a vector of 2-tuples for cells which depend on other cells, where the first
   element of the 2-tuple is the cell and the second is a cell on which it depends.
   The final vector may have multiple entries for a single cell if that cell has
   multiple dependencies."
  [{:keys [cells] :as wb-sheet-map}]
  (-> wb-sheet-map
      (assoc :dependencies
             (reduce (fn [accum {cell-sheet :sheet cell-label :label
                                 cell-formula :formula cell-references :references
                                 :as cell-map}]
                       (if cell-references
                         (concat
                          accum
                          (mapcat
                           (fn [{cr-type :type cr-sheet :sheet cr-label :label :as cell-reference}]
                             [[cell-map cell-reference]])
                           cell-references))
                         accum))
                     []
                     cells))))

(defn get-cell-dependencies
  [wb-map]
  (reduce
   (fn [accum [sheet-name wb-sheet-map]]
     (assoc accum
            sheet-name
            (get-cell-dependencies-for-sheet wb-sheet-map)))
   wb-map
   wb-map))

(defn get-cell-from-wb-map
  "Return the cell for a sheet and label, but without the :references key"
  ([cell-sheet cell-label wb-map-with-dependencies]
   (->> (get-in wb-map-with-dependencies [cell-sheet :cells])
        (some (fn [{:keys [sheet label] :as cell}]
                (when (and (= sheet cell-sheet)
                           (= label cell-label))
                  (dissoc cell :references)))))))

(defn add-self-dependencies-for-sheet
  "Add cells with formulas, but with no dependencies to the
   map as self-dependents"
  [wb-sheet-map-with-dependencies]
  (let [dependent-cells (:dependencies wb-sheet-map-with-dependencies)
        independent-cells (->> (:cells wb-sheet-map-with-dependencies)
                               (keep #(when
                                       (and (empty? (:references %))
                                            (some? (:formula %)))
                                        %)))]
    (reduce
     (fn [accum {:keys [sheet label] :as independent-cell}]
       ;; strictly speaking this should never be true
       (if (some #(= % independent-cell) accum)
         accum
         ;; add the cell to the map as having a formula and
         ;; depending on itself, so that we can force a recalc
         (conj accum [independent-cell {:sheet sheet :label "$$ROOT" :type :root}])))
     dependent-cells
     independent-cells)))

(defn add-self-dependencies
  [wb-map-with-dependencies]
  (reduce (fn [accum [sheet-name wb-map-with-dependencies]]
            (let [x (add-self-dependencies-for-sheet wb-map-with-dependencies)]
              (assoc-in accum [sheet-name :dependencies]
                        x)))
          wb-map-with-dependencies
          wb-map-with-dependencies))

(defn consolidate-dependencies-across-sheets [wb-map-with-dependencies]
  (mapcat (fn [[sheet-name {:keys [dependencies] :as wb-sheet-map}]]
            dependencies)
          wb-map-with-dependencies))

(defn add-graph
  ([wb-map-with-dependencies]
   (add-graph wb-map-with-dependencies true))
  ([wb-map-with-dependencies include-all-formula-cells?]
   (assoc wb-map-with-dependencies
          :graph
          (reduce (fn [accum [{cell-sheet :sheet cell-label :label :as cell}
                              {depends-sheet :sheet depends-label :label :as depends-on-cell}]]
                    (let [node-1 (str cell-sheet "!" cell-label)
                          node-2 (str depends-sheet "!" depends-label)
                          node-1-map (get-cell-from-wb-map cell-sheet cell-label wb-map-with-dependencies)
                          node-2-map (get-cell-from-wb-map depends-sheet depends-label wb-map-with-dependencies)]
                      (-> accum
                          (uber/add-nodes-with-attrs [node-1 node-1-map])
                          (uber/add-nodes-with-attrs [node-2 node-2-map])
                          (uber/add-edges [node-2 node-1]))))
                  (uber/digraph)
                  (cond-> wb-map-with-dependencies
                    include-all-formula-cells?
                    (add-self-dependencies)
                    true
                    (consolidate-dependencies-across-sheets))))))

(defn connect-disconnected-regions
  [{graph :graph :as wb-map-with-graph}]
  (let
   [sheet-root-node-labels (keep (fn [sheet-name]
                                   (when (string? sheet-name)
                                     [sheet-name (str sheet-name "!$$ROOT")]))
                                 (->> wb-map-with-graph
                                      (keys)))
    graph-with-roots (reduce
                      (fn [g [sheet-name sheet-root-node-label]]
                        (-> g
                            (uber/add-nodes-with-attrs
                             [sheet-root-node-label
                              {:label "$$ROOT" :sheet sheet-name :type :root}])
                            (uber/add-edges ["ROOT!$$ROOT" sheet-root-node-label])))
                      (-> graph
                          (uber/add-nodes-with-attrs
                           ["ROOT!$$ROOT"
                            {:label "$$ROOT" :sheet "ROOT" :type :root}]))
                      sheet-root-node-labels)]
    (->> (uber/nodes graph-with-roots)
         (reduce
          (fn [g n]
            (let [[cell-sheet cell-label] (str/split n #"\!")
                  id (uber/in-degree g n)]
              (if (and (not= n (str cell-sheet "!$$ROOT"))
                       (or
                        (= 0 id)
                        (and (= 1 id) (uber/find-edge g n n))))
                (let [node-with-map (get-cell-from-wb-map cell-sheet cell-label wb-map-with-graph)]
                  (-> g
                      (uber/add-nodes-with-attrs [n node-with-map])
                      (uber/add-edges [(str cell-sheet "!$$ROOT") n])))
                g)))
          graph-with-roots)
         (assoc wb-map-with-graph :graph))))

(defn get-recalc-node-sequence
  "Given an updated node at updated-node, return a sequence or other nodes that need
   to be recalculated, in the order they need to be recalculated"
  [updated-node {:keys [graph] :as wb-map}]
  (->>
   {:start-node updated-node}
   (alg/shortest-path graph)
   :depths
   (reduce (fn [accum [cell-name depth]]
             (update accum depth
                     (fnil conj [])
                     cell-name))
           (sorted-map))
   (reduce (fn [accum [_ cell-name]]
             (concat accum cell-name))
           [])))

(defn can-be-int? [v]
  (and (number? v)
       (= 0 (compare v (int v)))))

(defn eval-range [range-str wb-map]
  (let [expanded-range (expand-cell-range range-str wb-map)
        range-metadata (meta expanded-range)]
    (->> expanded-range
         (mapv (fn [{cell-sheet :sheet cell-label :label}]
                 (->> wb-map
                      (get-cell-from-wb-map cell-sheet cell-label)
                      (:value))))
         ((fn [interim-result]
            (cond (or (nil? interim-result)
                      (empty? interim-result))
                  nil
                  (= 1 (count interim-result))
                  (let [r (first interim-result)]
                    (if (can-be-int? r)
                      (int r)
                      r))
                  :else
                  interim-result)))
         ((fn [interim-result]
            (box/box interim-result range-metadata))))))

(defn substitute-ranges [unsubstituted-form]
  ;; TODO: the ns-resolve is not really needed. Could be replaced
  ;; with a pure 'graph/eval-range.
  (walk/postwalk
   (fn [form]
     (if (and (list? form)
              (= 'eval-range (first form)))
       (let [e (concat
                (cons (ns-resolve 'graph (first form)) (rest form))
                (list `*context*))]
         `(~@e))
       form))
   unsubstituted-form))

(defn construct-dynamic-range-for-range-operator
  [forms]
  (letfn [(->fn [f]
            (if (var? (eval f))
              (-> f eval deref)
              (eval f)))]
    (let [fs (mapv
              (fn [[f-name f-arg :as form]]
                ;; TODO: Originally here, the comparison was 
                ;; (= functions/fn-index (->fn f-name))
                ;; but this requires ns import of functions
                (cond (= 'functions/fn-index f-name)
                      (re-matches #"(.*!)?(.*)" (eval (cons 'functions/fn-index-reference (rest form))))
                      (= graph/eval-range (->fn f-name))
                      (re-matches #"(.*!)?(.*)" f-arg)
                      :else
                      (throw (IllegalArgumentException.
                              (str "NO MATCH. "
                                   "Form " (pr-str form)
                                   " FNAME:" f-name
                                   " M?:" (if (var? (eval f-name))
                                            (-> f-name eval deref)
                                            (eval f-name)))))))
              forms)
          fstr (str (some (fn [[_ sheet _]] (when (some? sheet) sheet)) fs)
                    (str/join ":" (map (fn [[_ _ label]] label) fs)))]
      `(eval-range ~fstr *context*))))

(defn substitute-dynamic-ranges
  [substituted-form]
  (walk/postwalk
   (fn [f]
     (if (and (list? f) (= 'functions/fn-range (first f)))
       (construct-dynamic-range-for-range-operator (-> f rest))
       f))
   substituted-form))

(defn results-equal?
  "Compare calculated result against the result as reported by 
   Excel, accommodating different types and some tolerance for
   rounding for numbers"
  [v1 v2]
  (let [v1-value (if (instance? Box v1) @v1 v1)
        v2-value (if (instance? Box v2) @v2 v2)
        numbers? (and (number? v1-value) (number? v2-value))
        compatible?
        (or
         (= (type v1-value) (type v2-value))
         numbers?)]
    (if compatible?
      (or (= 0 (compare v1-value v2-value))
          (and numbers? (< (math/abs (- v1-value v2-value)) 0.0000001M)))
      false
      #_(throw (Exception.
                (str "Invalid comparison values: "
                     v1 " [" (type v1) "], "
                     v2 " [" (type v2) "]"))))))

(comment
  (results-equal? 1 1)
  (results-equal? 1. 1)
  (results-equal? 1. "1")
  (results-equal? 1. 1.0000001)
  (results-equal? 1. 1.00000001)
  (results-equal? "A" "A")
  :end)

(defn substitute-indirection [form & [sheet-name]]
  (tap> {:loc substitute-indirection
         :f form
         :sheet-name sheet-name})
  (walk/postwalk
   (fn [f]
     (cond (and (list? f)
                (list? (first f))
                (= 'partial (ffirst f))
                (= 'functions/fn-indirect (-> f (first) (second))))
           (list 'eval-range f)
           (and false
                (list? f)
                (list? (first f))
                (= 'partial (ffirst f))
                (= 'functions/fn-offset (-> f (first) (second)))
                (list? (second f))
                (= 'eval-range (-> f (second) (first))))
           (concat (list (first f))
                   (list 'as-ref (second f))
                   (-> f (rest) (rest)))
           :else
           f))
   form))

(defn recalc-workbook
  "Recalculate a workbook's sheet. Standard assumption is that 
   a graph is available and that it's acyclic. If it's not acyclic
   this function will return no results. Don't supply a graph
   that's not a DAG"
  ([{:keys [graph] :as wb-map} & [sheet-name]]
   (reduce (fn [accum node]
             (let [[node {:keys [sheet formula value] :as attrs}]
                   (uber/node-with-attrs graph node)]
               (if (and formula (if sheet-name (= sheet-name sheet) true))
                 (let [formula-code (-> (str "=" formula)
                                        (parse/parse-to-tokens)
                                        (parse/nest-ast)
                                        (parse/wrap-ast)
                                        (ast/process-tree)
                                        (sh/parse-expression-tokens)
                                        (ast/unroll-for-code-form sheet))
                       formula-code-with-indirection (-> formula-code
                                                         (substitute-indirection sheet-name))
                       final-code (binding [*context* wb-map]
                                    (-> formula-code-with-indirection
                                        (substitute-ranges)
                                        (substitute-dynamic-ranges)))
                       calculated-result (binding [*context* wb-map]
                                           (-> final-code
                                               (eval)))]
                   #_(tap> {:base-code formula-code
                            :indirected-code formula-code-with-indirection
                            :final-code final-code})
                   (conj
                    accum
                    [node (results-equal? value calculated-result)
                     formula value final-code
                     (if (instance? Box calculated-result) @calculated-result calculated-result)]))
                 accum)))
           []
           (alg/topsort graph))))

(defn simplify-results
  [recalc-results]
  (->> recalc-results
       (mapv
        (fn [[cell-name vals-match? formula-text
              excel-value clj-code calc-value]]
          {:cell cell-name
           :formula formula-text
           :clj-code clj-code
           :excel-value excel-value
           :clj-value calc-value}))))

(comment

  {:vlaaad.reveal/command '(clear-output)}

  (-> "TEST-cyclic.xlsx"
      (explain-workbook "Sheet1"))

  (-> "TEST-cyclic.xlsx"
      (explain-workbook "Sheet1")
      (get-cell-dependencies))

  (-> "TEST-cyclic.xlsx"
      (explain-workbook "Sheet1")
      (get-cell-dependencies)
      (add-self-dependencies))

  (-> "TEST-cyclic.xlsx"
      (explain-workbook "Sheet1")
      (get-cell-dependencies)
      (add-graph false))

  (-> "TEST-cyclic.xlsx"
      (explain-workbook "Sheet1")
      (get-cell-dependencies)
      (add-graph)
      (connect-disconnected-regions)
      (:graph)
      (uber/viz-graph))

  (-> "SIMPLE-1.xlsx"
      (explain-workbook)
      (get-cell-dependencies)
      (add-graph)
      (connect-disconnected-regions)
      (:graph)
      (uber/viz-graph {:save {:filename "./assets/DAG3.png"
                              :format :png}
                       :auto-label true}))

  (-> "SIMPLE-1.xlsx"
      (explain-workbook)
      (get-cell-dependencies)
      (add-graph)
      (connect-disconnected-regions)
      (recalc-workbook "Scores")
      (simplify-results))

  (def WB-MAP
    (-> "TEST-cyclic.xlsx"
        (explain-workbook "Sheet1")
        (get-cell-dependencies)
        (add-graph)
        (connect-disconnected-regions)))

  (uber/viz-graph (:graph WB-MAP))

  (ubergraph.alg/connected-components (:graph WB-MAP))

  (uber/in-degree (:graph WB-MAP) "Sheet1!C1")
  (uber/edges (:graph WB-MAP))
  (uber/nodes (:graph WB-MAP))
  (ubergraph.alg/shortest-path (:graph WB-MAP) {:start-node "Sheet1!B1"})
  (map (partial uber/edge-with-attrs
                (:graph WB-MAP))
       (uber/edges (:graph WB-MAP)))
  (ubergraph.alg/scc (:graph WB-MAP))
  (uber/node-with-attrs (:graph WB-MAP) "Sheet1!B4")
  (uber/node-with-attrs (:graph WB-MAP) "Sheet1!$$ROOT")
  (ubergraph.alg/connect (:graph WB-MAP))
  (uber/viz-graph (ubergraph.alg/connect (:graph WB-MAP)))
  (uber/nodes (:graph WB-MAP))
  (uber/pprint (:graph WB-MAP))
  (uber/viz-graph (:graph WB-MAP))
  (uber/node-with-attrs (:graph WB-MAP) "Sheet3!B3")
  (ubergraph.alg/dag? (:graph WB-MAP))

  (expand-cell-range "Sheet2!B3:D3" (:named-ranges WB-MAP))
  (expand-cell-range "Sheet2!BONUS" (:named-ranges WB-MAP))
  (expand-cell-range "Sheet2!B2" (:named-ranges WB-MAP))

  (eval-range "Sheet2!C2:C4" WB-MAP)
  (eval-range "Sheet2!ALLOWEDTOTAL" WB-MAP)
  (eval-range "Sheet2!J4:J6" WB-MAP)

  (binding [*context* WB-MAP]
    (-> (substitute-ranges
         '(if (< (eval-range "Sheet2!G5") (eval-range "Sheet2!ALLOWEDTOTAL")) (str "YES") (str "NO")))
        (eval)))

  (binding [*context* WB-MAP]
    (-> '(functions/fn-counta (eval-range "Sheet2!EMPLOYEES"))
        (substitute-ranges)
        (eval)))

  (binding [*context* WB-MAP]
    (eval-range "Sheet2!EMPLOYEES" *context*))

  (alg/topsort (:graph WB-MAP))

  (recalc-workbook WB-MAP "Sheet3")

  (keep (fn [[cell-label match? cell-formula cell-value cell-code calculated-value :as calc]]
          (when-not match?
            calc))
        (recalc-workbook WB-MAP "Sheet2"))

  (uber/node-with-attrs (:graph WB-MAP) "Sheet3!B3")

  (get-recalc-node-sequence "Sheet2!A2" WB-MAP)

  (def G
    (-> "TEST1.xlsx"
        (explain-workbook "Sheet2")
        (get-cell-dependencies)
        (add-graph)
        (connect-disconnected-regions)
        :graph))

  (uber/pprint G)
  (uber/viz-graph G)
  (uber/node-with-attrs G "Sheet2!C4")

  (-> "TEST1.xlsx"
      (explain-workbook "Sheet2")
      (get-cell-dependencies)
      (add-graph)
      (connect-disconnected-regions)
      :graph
      (uber/viz-graph))

  (reduce (fn [accum node]
            (conj
             accum
             (let [[node {:keys [formula value] :as attrs}]
                   (uber/node-with-attrs G node)]
               [node formula value])))
          []
          (alg/topsort G))

  (uber/pprint G)
  (uber/find-edges G :c :e)
  (uber/has-node? G :a)
  (uber/node-with-attrs G :c)
  (uber/viz-graph G)
  (alg/topsort G)
  (alg/shortest-path G {:start-node "Sheet2!B9" :traverse true})
  (alg/paths->graph (alg/shortest-path G {:start-node "Sheet2!B9"}))
  (-> (alg/paths->graph (alg/shortest-path G {:start-node "Sheet2!B9"})) (uber/viz-graph))
  (alg/pprint-path (alg/all-destinations (alg/shortest-path G {:start-node "Sheet2!B3"})))

  (alg/nodes-in-path
   (alg/shortest-path G {:start-node "Sheet2!B3" :traverse false}))


  :end)
