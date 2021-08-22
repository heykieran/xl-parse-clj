(ns core
  (:require
   [xlparse :as parse]
   [shunting :as sh]
   [shunting-xl :as shxl]
   [ast-processing :as ast]
   [clojure.walk :as walk]
   [excel :as excel]
   [functions]))

(comment
  (-> "=(1+2+3)*(10+20+30)/(40+50+60)"
      (parse/parse-to-tokens)
      (parse/nest-ast))
  (-> "=1+2+3"
      (parse/parse-to-tokens)
      #_(parse/nest-ast))
  (shxl/parse-simple-expression-tokens
   [{:value "1", :type :Operand, :sub-type :Number}
    {:value "+", :type :OperatorInfix, :sub-type :Math}
    {:value "2", :type :Operand, :sub-type :Number}
    {:value "+", :type :OperatorInfix, :sub-type :Math}
    {:value "3", :type :Operand, :sub-type :Number}])

  (shxl/parse-simple-expression-tokens
   [{:value "10", :type :Operand, :sub-type :Number}
    {:value "+", :type :OperatorInfix, :sub-type :Math}
    {:value "20", :type :Operand, :sub-type :Number}
    {:value "+", :type :OperatorInfix, :sub-type :Math}
    {:value "30", :type :Operand, :sub-type :Number}])

  (shxl/parse-expression-tokens
   [[{:value "1", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "2", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "3", :type :Operand, :sub-type :Number}]
    {:value "*", :type :OperatorInfix, :sub-type :Math}
    [{:value "10", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "20", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "30", :type :Operand, :sub-type :Number}]])
  :end)

(def ast-raw
  (-> "=(1+2+3)*(10+20+30)/(40+50+60)"
      (parse/parse-to-tokens)
      (parse/nest-ast)))

(def ast
  {0
   {:value ""
    :type :Subexpression
    :sub-type :Start
    1
    {:value ""
     :type :Subexpression
     :sub-type :Start
     2 {:value "1", :type :Operand, :sub-type :Number}
     3 {:value "+", :type :OperatorInfix, :sub-type :Math}
     4 {:value "2", :type :Operand, :sub-type :Number}
     5 {:value "+", :type :OperatorInfix, :sub-type :Math}
     6 {:value "3", :type :Operand, :sub-type :Number}}
    8 {:value "*", :type :OperatorInfix, :sub-type :Math}
    9
    {:value ""
     :type :Subexpression
     :sub-type :Start
     10 {:value "10", :type :Operand, :sub-type :Number}
     11 {:value "+", :type :OperatorInfix, :sub-type :Math}
     12 {:value "20", :type :Operand, :sub-type :Number}
     13 {:value "+", :type :OperatorInfix, :sub-type :Math}
     14 {:value "30", :type :Operand, :sub-type :Number}}
    16 {:value "/", :type :OperatorInfix, :sub-type :Math}
    17
    {:value ""
     :type :Subexpression
     :sub-type :Start
     18 {:value "40", :type :Operand, :sub-type :Number}
     19 {:value "+", :type :OperatorInfix, :sub-type :Math}
     20 {:value "50", :type :Operand, :sub-type :Number}
     21 {:value "+", :type :OperatorInfix, :sub-type :Math}
     22 {:value "60", :type :Operand, :sub-type :Number}}}})

(defn should-simplify? [form]
  (and
   (map? form)
   (identical?
    :Subexpression
    (:type form))))

(defn expressions->vector [exps]
  (if (map? exps)
    (vector exps)
    (into [] exps)))

(defn add-expression [form & [debug]]
  (assoc
   form
   :clj
   (expressions->vector
    (shxl/parse-simple-expression-tokens
     (keep
      (fn [[k v]]
        (when
         (number? k)
          v))
      form)
     debug))))

(defn filter-expression [form]
  (apply dissoc form
         (filter #(not= % :clj) (keys form))))

(defn as-code [ast & {:keys [debug]}]
  (walk/postwalk
   (fn [form]
     (if (should-simplify? form)
       (-> form
           (add-expression debug)
           (filter-expression))
       form))
   ast))

(defn validate [r-map]
  (println "Check" (:value r-map) "=" (:result r-map) "for" (:formula r-map))
  (assoc r-map :ok? (= (:value r-map)
                       (:result r-map))))

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
  (as-code
   (-> "=(1+2+3)*(10+20+30)/(40+50+60)"
       (parse/parse-to-tokens)
       (parse/nest-ast)
       (parse/wrap-ast)))

  (as-code
   (-> "=(1+2+3)+4"
       (parse/parse-to-tokens)
       (parse/nest-ast)
       (parse/wrap-ast)))

  (as-code
   (-> "=(1+2+3)"
       (parse/parse-to-tokens)
       (parse/nest-ast)
       (parse/wrap-ast)))

  (as-code
   (-> "=max(1,max(1,(2),4))"
       (parse/parse-to-tokens)
       (parse/nest-ast)
       (parse/wrap-ast))
   :debug false)

  (as-code
   (-> "=1+max(1,2)"
       (parse/parse-to-tokens)
       (parse/nest-ast)
       (parse/wrap-ast))
   :debug false)

  (-> "=max(1,max(1,(2),4))"
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast))

  (as-code
   (-> "=max(1,max(1,(2),4))"
       (parse/parse-to-tokens)
       (parse/nest-ast)
       (parse/wrap-ast)))

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

  (-> "=COUNTA(A2:A4)"
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

