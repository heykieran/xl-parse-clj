(ns core
  (:require
   [xlparse :as parse]
   [shunting :as sh]
   [ast-processing :as ast]
   [excel :as excel]
   [functions :as functions]))

(comment
  (-> "=(1+2+3)*(10+20+30)/(40+50+60)"
      (parse/parse-to-tokens)
      (parse/nest-ast))
  (-> "=1+2+3"
      (parse/parse-to-tokens)
      (parse/nest-ast))
  :end)

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

(defn run-tests []
  (->> (excel/extract-test-formulas "TEST1.xlsx" "Sheet1")
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
        [])))


(comment

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

(comment

  (functions/fn-sumif [118.0 229.0 340.0] (str ">200"))
  (functions/fn-sumif (graph/eval-range "Sheet2!J4:J6" WB-MAP) 
                      (graph/eval-range "Sheet2!E4" WB-MAP))
  
  (graph/eval-range "Sheet2!G4" WB-MAP)

  (filter #(> % 200) [118.0 229.0 340.0])

  (require '[clojure.tools.analyzer.jvm :as ana.jvm])
  (require '[clojure.tools.analyzer.passes.jvm.emit-form :as e])
  (require '[graph :as graph])

  (ana.jvm/analyze '(> 200.0))
  (ana.jvm/analyze '(graph/eval-range "E1"))

  (ana.jvm/analyze '(> 200.0) {})
  (e/emit-form (ana.jvm/analyze '(> 200.0)))

  (-> "=IF(X>200,1,0)" #_"=$CURRENT>200"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form))
  
  :end)

