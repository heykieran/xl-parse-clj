(ns graph
  (:require [ubergraph.core :as uber]
            [ubergraph.alg :as alg]
            [dk.ative.docjure.spreadsheet :as dk]
            [xlparse :as parse]
            [excel :as excel]
            [ast-processing :as ast]
            [shunting :as sh]
            [core :as core]
            [clojure.walk :as walk])
  (:import
   [java.time LocalDate LocalTime LocalDateTime]
   [java.time.format DateTimeFormatter]
   [org.apache.poi.ss SpreadsheetVersion]
   [org.apache.poi.ss.util AreaReference CellReference]
   [org.apache.poi.ss.usermodel CellType DateUtil]))

(declare ^:dynamic *context*)

(defn run-tests []
  (->> (excel/extract-test-formulas "TEST1.xlsx" "Sheet2")
       (reduce
        (fn [accum {:keys [type formula format address row column value
                           :calc-date-value :excel-date-value] :as fcell-info}]
          (conj accum
                (let [clj-exp (-> (str "=" formula)
                                  (parse/parse-to-tokens)
                                  (parse/nest-ast)
                                  (parse/wrap-ast)
                                  (ast/process-tree)
                                  (sh/parse-expression-tokens)
                                  (ast/unroll-for-code-form))]
                  (-> fcell-info
                      (assoc
                       :clj
                       clj-exp
                       :result
                       (-> (eval clj-exp)
                           (core/convert-result)))
                      (core/validate)))))
        [])))

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
   (if-let [named-range (get-named-range-description-map named-ranges cell-range-str)]
     (:references named-range)
     (let [aref (AreaReference. cell-range-str SpreadsheetVersion/EXCEL2007)]
       (mapv (fn [ref-cell]
               (let [[sheet-name row-name col-name]
                     (.getCellRefParts ref-cell)]
                 {:sheet sheet-name
                  :label (str col-name row-name)
                  :type :general}))
             (.getAllReferencedCells aref))))))

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
  ([wb-name]
   (explain-workbook wb-name "Sheet2"))
  ([wb-name sheet-name]
   (let [wb-as-resource (dk/load-workbook-from-resource wb-name)]
     (explain-cells-in-sheet wb-as-resource sheet-name))))

(defn get-cell-dependencies
  "Returns a vector of 2-tuples for cells which depend on other cells, where the first
   element of the 2-tuple is the cell and the second is a cell on which it depends.
   The final vector may have multiple entries for a single cell if that cell has
   multiple dependencies."
  [{:keys [cells] :as wb-map}]
  (-> wb-map
      (assoc :dependencies
             (reduce (fn [accum {cell-label :label cell-formula :formula cell-references :references :as cell-map}]
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

(defn get-cell-from-wb-map
  "Return the cell for a sheet and label, but without the :references key"
  [cell-sheet cell-label {:keys [cells] :as wb-map-with-dependencies}]
  (some (fn [{:keys [sheet label] :as cell}]
          (when (and (= sheet cell-sheet)
                     (= label cell-label))
            (dissoc cell :references)))
        cells))

(defn add-graph [{:keys [dependencies] :as wb-map-with-dependencies}]
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
                 dependencies)))

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
  (->> (expand-cell-range range-str named-ranges)
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
                interim-result)))))

(defn substitute-ranges [unsubstituted-form]
  (walk/postwalk
   (fn [form]
     (if (and (list? form) 
              (= 'eval-range (first form)))
       (let [e (concat 
                (cons (resolve (first form)) (rest form)) 
                (list `*context*))]
         `(~@e))
       form))
   unsubstituted-form))

(defn recalc-workbook [{:keys [graph] :as wb-map} sheet-name]
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
          (alg/topsort graph)))

(comment

  {:vlaaad.reveal/command '(clear-output)}

  (explain-workbook "TEST1.xlsx" "Sheet2")

  (-> "TEST1.xlsx"
       (explain-workbook "Sheet2")
       (get-cell-dependencies))

  (-> "TEST1.xlsx"
       (explain-workbook)
       (get-cell-dependencies)
       (add-graph))

  (def WB-MAP
    (-> "TEST1.xlsx"
        (explain-workbook "Sheet2")
        (get-cell-dependencies)
        (add-graph)))

  (expand-cell-range "Sheet2!B3:D3" (:named-ranges WB-MAP))
  (expand-cell-range "Sheet2!BONUS" (:named-ranges WB-MAP))
  (expand-cell-range "Sheet2!B2" (:named-ranges WB-MAP))

  (eval-range "Sheet2!C2:C4" WB-MAP)
  (eval-range "Sheet2!ALLOWEDTOTAL" WB-MAP)

  (binding [*context* WB-MAP]
    (-> (substitute-ranges
         '(if (< (eval-range "Sheet2!E5") (eval-range "Sheet2!ALLOWEDTOTAL")) (str "YES") (str "NO")))
        (eval)))

  (binding [*context* WB-MAP]
    (-> '(functions/fn-counta (eval-range "Sheet2!EMPLOYEES"))
        (substitute-ranges)
        (eval)))

  (binding [*context* WB-MAP]
    (eval-range "Sheet2!EMPLOYEES" *context*))
  
  (recalc-workbook WB-MAP "Sheet2")

  (keep (fn [[cell-label match? cell-formula cell-value cell-code calculated-value :as calc]]
          (when-not match?
            calc))
        (recalc-workbook WB-MAP "Sheet2"))

  (if (< (eval-range "Sheet2!E5") (eval-range "Sheet2!ALLOWEDTOTAL")) (str "YES") (str "NO"))

  (uber/node-with-attrs (:graph WB-MAP) "Sheet2!C4")

  (get-recalc-node-sequence "Sheet2!A2" WB-MAP)

  (def G
    (-> "TEST1.xlsx"
        (explain-workbook "Sheet2")
        (get-cell-dependencies)
        (add-graph)
        :graph))

  (uber/pprint G)
  (uber/viz-graph G)
  (uber/node-with-attrs G "Sheet2!A2")

  (-> "TEST1.xlsx"
      (explain-workbook "Sheet2")
      (get-cell-dependencies)
      (add-graph)
      :graph
      (uber/viz-graph))

  (reduce (fn[accum node]
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
   (reduce (fn[accum [cell-name depth]]
             (update accum depth
                     (fnil conj [])
                     cell-name))
           (sorted-map))
   (reduce (fn[accum [depth cell-name]]
             (concat accum cell-name))
           []))

  :end
  )