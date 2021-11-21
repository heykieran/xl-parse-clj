(ns graph
  (:require
   [clojure.string :as str]
   [ubergraph.core :as uber]
   [ubergraph.alg :as alg]
   [dk.ative.docjure.spreadsheet :as dk]
   [xlparse :as parse]
   [excel :as excel]
   [ast-processing :as ast]
   [shunting :as sh]
   [clojure.walk :as walk])
  (:import
   [org.apache.poi.ss SpreadsheetVersion]
   [org.apache.poi.ss.util AreaReference]
   [org.apache.poi.ss.usermodel CellType DateUtil]))

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
    {:single? (.isSingleCell aref)
     :column? (.isWholeColumnReference aref)
     :sheet-name cell-sheet-name
     :tl-name (str cell-col-name cell-row-name)
     :tl-coord [f-row f-col]
     :cols (inc (apply - (sort > [l-col f-col])))
     :rows (inc (apply - (sort > [l-row f-row])))}))

(defn get-named-range-description-map [named-ranges cell-name]
  (when (and named-ranges (-> cell-name (AreaReference. SpreadsheetVersion/EXCEL2007) (.isSingleCell)))
    (some (fn [{:keys [name sheet] :as named-range}]
            (when (= cell-name (str sheet "!" name))
              named-range))
          named-ranges)))

(defn expand-cell-range
  ([cell-range-str]
   (expand-cell-range cell-range-str nil))
  ([cell-range-str named-ranges]
   (letfn [(cell-info [ref-cell]
             (let [[sheet-name row-name col-name]
                   (.getCellRefParts ref-cell)]
               {:sheet sheet-name
                :label (str col-name row-name)
                :type :general}))]
     (if-let [named-range (get-named-range-description-map named-ranges cell-range-str)]
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

(defn explain-named-ranges-in-workbook [wb-as-resource]
  (->>
   wb-as-resource
   (.getAllNames)
   (map (fn [n]
          {:name (.getNameName n)
           :sheet (.getSheetName n)
           :references (expand-cell-range (.getRefersToFormula n))}))))

(defn explain-cells-in-sheet [wb-as-resource sheet-name]
  (let [named-ranges (explain-named-ranges-in-workbook wb-as-resource)]
    (->> wb-as-resource
         (dk/select-sheet sheet-name)
         dk/cell-seq
         (explain-sheet-cells sheet-name)
         (add-references sheet-name named-ranges)
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
     (reduce
      (fn [accum sheet-name]
        (assoc accum sheet-name
               (explain-cells-in-sheet wb-as-resource sheet-name)))
      {}
      sheet-names))))

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
   (->> wb-map-with-dependencies
        (get cell-sheet)
        (:cells)
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
              (tap> {:loc add-self-dependencies
                      :x x})
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
   (add-graph wb-map-with-dependencies false))
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
                    (consolidate-dependencies-across-sheets))
                  #_(if include-all-formula-cells?
                    (add-self-dependencies-for-sheet wb-map-with-dependencies)
                    (consolidate-dependencies-across-sheets wb-map-with-dependencies))))))

(defn connect-disconnected-regions 
  [{graph :graph :as wb-map-with-graph}]
  (let 
   [sheet-root-node-labels (keep (fn [sheet-name]
                                   (when (string? sheet-name)
                                     (str sheet-name "!$$ROOT")))
                                 (->> wb-map-with-graph
                                      (keys)))
    graph-with-roots (reduce
                      (fn [g sheet-root-node-label]
                        (-> g 
                            (uber/add-nodes-with-attrs [sheet-root-node-label {}])
                            (uber/add-edges ["ROOT!$$ROOT" sheet-root-node-label])))
                      (-> graph
                          (uber/add-nodes-with-attrs ["ROOT!$$ROOT" {}]))
                      sheet-root-node-labels)]
    (->> (uber/nodes graph-with-roots)
         (reduce
          (fn [g n]
            (let [[cell-sheet cell-label] (str/split n #"\!")
                  id (uber/in-degree g n)]
              (tap> {:loc connect-disconnected-regions
                     :n n
                     :id id
                     :s cell-sheet
                     :l cell-label
                     :e (uber/find-edge g n n)})
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

(defn eval-range [range-str {:keys [named-ranges] :as wb-map}]
  (let [expanded-range (expand-cell-range range-str named-ranges)
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
                  (first interim-result)
                  :else
                  interim-result)))
         ((fn[interim-result]
            (if (instance? clojure.lang.IObj interim-result)
              (with-meta interim-result range-metadata)
              interim-result))))))

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

(defn recalc-workbook
  "Recalculate a workbook's sheet. Standard assumption is that 
   a graph is available and that it's acyclic. If it's not acyclic
   this function will return no results. However, by setting 
   recalc-all?, a calculation can be forced over all the nodes 
   that include formulae, by using a different ordering algorithmn
   that topological sort."
  ([wb-map sheet-name]
   (recalc-workbook wb-map sheet-name false))
  ([{:keys [graph] :as wb-map} sheet-name recalc-all?]
   (reduce (fn [accum node]
             (let [[node {:keys [sheet formula value] :as attrs}]
                   (uber/node-with-attrs graph node)]
               (if formula
                 (let [formula-code (-> (str "=" formula)
                                        (parse/parse-to-tokens)
                                        (parse/nest-ast)
                                        (parse/wrap-ast)
                                        (ast/process-tree)
                                        (sh/parse-expression-tokens)
                                        (ast/unroll-for-code-form sheet-name))
                       calculated-result (binding [*context* wb-map]
                                           (-> formula-code
                                               (substitute-ranges)
                                               (eval)))]
                   (conj
                    accum
                    [node (= value calculated-result) formula value formula-code calculated-result]))
                 accum)))
           []
           (if recalc-all?
             (reverse (alg/post-traverse graph))
             (alg/topsort graph)))))

(comment

  {:vlaaad.reveal/command '(clear-output)}

  (explain-workbook "TEST-cyclic.xlsx" "Sheet1")

  (-> "TEST-cyclic.xlsx"
      (explain-workbook)
      (get-cell-dependencies)
      (add-graph true)
      (connect-disconnected-regions)
      (:graph)
      (uber/viz-graph))
  
  (-> "TEST-cyclic.xlsx"
      (explain-workbook "Sheet1")
      (get "Sheet1")
      (get-cell-dependencies))
  
  (-> "TEST-cyclic.xlsx"
      (explain-workbook "Sheet1")
      (get "Sheet1")
      (get-cell-dependencies)
      (add-self-dependencies-for-sheet))

  (-> "TEST-cyclic.xlsx"
      (explain-workbook "Sheet1")
      (get "Sheet1")
      (get-cell-dependencies)
      (add-graph))

  (def WB-MAP
    (-> "TEST-cyclic.xlsx"
        (explain-workbook "Sheet1")
        (get "Sheet1")
        (get-cell-dependencies)
        (add-graph true)
        (connect-disconnected-regions)))
  
  (uber/viz-graph (:graph WB-MAP))

  (ubergraph.alg/connected-components (:graph WB-MAP))
  
  (-> (let [g (:graph WB-MAP)
        r "Sheet1!$$ROOT"]
    (reduce (fn [g n]
              (let [id (uber/in-degree g n)]
                (if (and (not= n r)
                         (or
                          (= 0 id)
                          (and (= 1 id) (uber/find-edge g n n))))
                  (let [[cell-sheet cell-label] (clojure.string/split n #"\!")
                        node-with-map (get-cell-from-wb-map cell-sheet cell-label WB-MAP)]
                    (-> g 
                        (uber/add-nodes-with-attrs [n node-with-map])
                        (uber/add-edges [r n])))
                  g)))
            (-> g (uber/add-nodes-with-attrs [r {}]))
            (uber/nodes g)))
      (uber/viz-graph))
  
  (let [node-1 (str cell-sheet "!" cell-label)
        node-2 (str depends-sheet "!" depends-label)
        node-1-map (get-cell-from-wb-map cell-sheet cell-label wb-map-with-dependencies)
        node-2-map (get-cell-from-wb-map depends-sheet depends-label wb-map-with-dependencies)]
    (-> accum
        (uber/add-nodes-with-attrs [node-1 node-1-map])
        (uber/add-nodes-with-attrs [node-2 node-2-map])
        (uber/add-edges [node-2 node-1])))
  
  (uber/in-degree (:graph WB-MAP) "Sheet1!C1")
  (uber/edges (:graph WB-MAP) "Sheet1!C1")
  (ubergraph.alg/shortest-path (:graph WB-MAP) {:start-node "Sheet1!B1"})
  (map (partial uber/edge-with-attrs 
                (:graph WB-MAP)) 
       (uber/edges (:graph WB-MAP)))
  (ubergraph.alg/scc (:graph WB-MAP))
  (uber/node-with-attrs (:graph WB-MAP) "Sheet1!B4")
  (ubergraph.alg/connect (:graph WB-MAP))
  (uber/viz-graph (ubergraph.alg/connect (:graph WB-MAP)))
  (uber/nodes (:graph WB-MAP))
  (uber/pprint (:graph WB-MAP))
  (uber/viz-graph (:graph WB-MAP))
  (uber/node-with-attrs (:graph WB-MAP) "Sheet3!B3")


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
        (explain-workbook-sheet "Sheet2")
        (get-cell-dependencies)
        (add-graph)
        :graph))

  (uber/pprint G)
  (uber/viz-graph G)
  (uber/node-with-attrs G "Sheet2!A2")

  (-> "TEST1.xlsx"
      (explain-workbook-sheet "Sheet2")
      (get-cell-dependencies)
      (add-graph)
      :graph
      (uber/viz-graph))

  (reduce (fn [accum node]
            (conj
             accum
             (let [[node {:keys [formula value] :as attrs}] (uber/node-with-attrs G node)]
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

  (->>
   {:start-node "Sheet2!B3"}
   (alg/shortest-path G)
   :depths
   (reduce (fn [accum [cell-name depth]]
             (update accum depth
                     (fnil conj [])
                     cell-name))
           (sorted-map))
   (reduce (fn [accum [depth cell-name]]
             (concat accum cell-name))
           []))

  :end
  )