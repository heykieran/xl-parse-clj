(ns core
  (:require
   [xlparse :as parse]
   [shunting :as sh]
   [ast-processing :as ast]
   [excel :as excel]
   [graph :as graph]
   [functions]))

(defn compare-ok? [{:keys [value result] :as r-map}]
  (if (and (number? result) (number? value))
    (< -1e-8 (- value result) 1e-8)
    (= value result)))

(defn validate [r-map]
  (println "Check" (:value r-map) "=" (:result r-map) "for" (:formula r-map))
  (assoc r-map :ok? (compare-ok? r-map)))

(defn convert-result [v]
  (if (number? v)
    (double v)
    v))

(defn run-tests
  ([]
   (run-tests "TEST1.xlsx" "Sheet1"))
  ([workbook-name sheet-name]
   (->> (excel/extract-test-formulas workbook-name sheet-name)
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
                            (convert-result)))
                       (validate)))))
         []))))

(defn test-recalc-worksheet
  ([workbook-name worksheet-name]
   (mapv (fn [[node match? formula value formula-code calculated-result]]
           {:cell node
            :match? match?
            :result calculated-result
            :excel-value value
            :formula formula})
         (->
          (graph/explain-workbook workbook-name worksheet-name)
          (graph/get-cell-dependencies)
          (graph/add-graph)
          (graph/connect-disconnected-regions)
          (graph/recalc-workbook worksheet-name)))))

(comment

  (run-tests)
  (run-tests "TEST1.xlsx" "Sheet1")
  (test-recalc-worksheet "TEST1.xlsx" "Sheet1")
  (test-recalc-worksheet "TEST1.xlsx" "Sheet3")
  (test-recalc-worksheet "TEST-cyclic.xlsx" "Sheet3")

  (-> "=max(1,max(1,(2),4))"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast))

  (-> "=max(1,max(1,(2+3),4))"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast))

  (-> "=max(1,-2+3)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast))

  (-> "=max(1,-2+3)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree))

  (-> "=max(1,-2+3)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens))

  (-> "=max(1,-2+3,sin(4),max(5+6,7+8))"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form)
      (eval))

  (-> "=\"A\" & \"B\""
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form)
      (eval))

  (-> "=max(1,2)+max(3,4)-sin(5)*$A$4"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form))

  (-> "=if(1=2,-3+4,4+5)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form))

  (-> "=3 * 2%"
      (parse/parse-to-tokens))

  (-> "= 1 * 2% + 3 "
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form)
      (eval))

  (-> "= sin(100)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form)
      (eval))

  (-> "= 1% + 2% + 3%"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form)
      (eval))

  (-> "= 1% / 2 + 3"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form)
      (eval))

  (-> "=ABS(-200.3)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form)
      (eval))

  (-> "=OR(1,2,3)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form))

  (-> "=MAX(1,2,3,4=4,4<5)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      #_(ast/unroll-for-code-form))

  (-> "=YEARFRAC(\"2001/01/25\",\"2001/09/27\")"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form))

  (-> "=SUMIF(J4:J6,\">200\")"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form))

  (-> "=SUMIF(J4:J6,E1)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form))

  (-> "=IF(X>200,1,0)"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form))

  (->> (run-tests)
       (filter #(false? (:ok? %))))

  (run-tests)

  :end)
