(ns core
  (:require
   [xlparse :as parse]
   [shunting :as sh]
   [ast-processing :as ast]
   [clojure.walk :as walk]
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
  
  (->> (run-tests)
       (filter #(false? (:ok? %))))

  (run-tests)

  (boolean 1.0)
  (not (boolean 0))
  (functions/fn-not (if (not= 1.0 1.0) 1.0 0.0))
  (some #(true? (boolean %)) '(1.0 2.0 3.0))

  :end)

