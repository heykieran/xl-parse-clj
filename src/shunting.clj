(ns shunting
  (:require [clojure.string :as str]
            [clojure.walk :as walk]
            [functions :as fns]))

(def OPERATORS-PRECEDENCE
  [{:name :unary-plus :s "+" :f '+ :c :prefix :a 1 :e [:OperatorPrefix nil]}
   {:name :unary-minus :s "-" :f '- :c :prefix :a 1 :e [:OperatorPrefix nil]}

   {:name :unary-prcnt :s "%" :f `fns/prcnt :c :postfix :a 1 :e [:OperatorPostfix nil]}

   {:name :binary-mult :s "*" :f '* :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-div :s "/" :f '/  :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-plus :s "+" :f '+ :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-minus :s "-" :f '- :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-exp :s "^" :f '** :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-concat :s "&" :f 'str :c :infix :a 2 :e [:OperatorInfix :Concatenation]} ; what should this do? There are coercions to text in Excel

   {:name :compare-eq :s "=" :f '= :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-gt :s ">" :f '> :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-lt :s "<" :f '< :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-gte :s ">=" :f '>= :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-lte :s "<=" :f '<= :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-neq :s "<>" :f 'not= :c :infix :a 2 :e [:OperatorInfix :Logical]}

   {:name :abs :s "abs" :f `fns/abs :c :args :a 1 :e [:Function :Start]}
   {:name :sin :s "sin" :f 'Math/sin :c :args :a 1 :e [:Function :Start]}
   {:name :true :s "true" :f `fns/fn-true :c :args :a 0 :e [:Function :Start]}
   {:name :false :s "false" :f `fns/fn-false :c :args :a 0 :e [:Function :Start]}
   {:name :and :s "and" :f `fns/fn-and :c :args :a :all :e [:Function :Start]}
   {:name :or :s "or" :f `fns/fn-or :c :args :a :all :e [:Function :Start]}
   {:name :not :s "not" :f `fns/fn-not :c :args :a 1 :e [:Function :Start]}
   {:name :max :s "max" :f 'max :c :args :a :all :e [:Function :Start]}
   {:name :min :s "min" :f 'min :c :args :a :all :e [:Function :Start]}
   {:name :pi :s "pi" :f `fns/pi :c :args :a 0 :e [:Function :Start]}
   {:name :sum :s "sum" :f `fns/sum :c :args :a :all :e [:Function :Start]}
   {:name :average :s "average" :f `fns/average :c :args :a :all :e [:Function :Start]}
   {:name :count :s "count" :f `fns/fn-count :c :args :a :all :e [:Function :Start]}
   {:name :count :s "counta" :f `fns/fn-counta :c :args :a :all :e [:Function :Start]}
   {:name :now :s "now" :f `fns/fn-now :c :args :a 0 :e [:Function :Start]}
   {:name :date :s "date" :f `fns/fn-date :c :args :a :all :e [:Function :Start]}
   {:name :days :ext true :s "_xlfn.days" :f `fns/fn-days :c :args :a :all :e [:Function :Start]}
   {:name :datevalue :s "datevalue" :f `fns/fn-datevalue :c :args :a 1 :e [:Function :Start]}
   {:name :yearfrac :s "yearfrac" :f `fns/fn-yearfrac :c :args :a :all :e [:Function :Start]}

   {:name :if :s "if" :f 'if :c :args :a 3 :e [:Function :Start]}])

(defn =ci [a1 a2]
  (= (some-> a1 str/lower-case)
     (some-> a2 str/lower-case)))

(defn get-args-count [type sub-type value]
  (some (fn [{o-value :s [e-type e-sub-type] :e num-args :a}]
          (when (and (=ci o-value value)
                     (= e-type type)
                     (= e-sub-type sub-type))
            num-args))
        OPERATORS-PRECEDENCE))

(def STACK-DEBUG false)

(defn println-dbg [msg & [v]]
  (when STACK-DEBUG
    (print (str msg " "))
    (when v
      (println
       (walk/postwalk
        (fn [f]
          (if (map? f)
            (:value f)
            f))
        (if (sequential? v) v [v]))))))

(defn get-operator-fn
  ([operator-str type sub-type]
   (let [o-fn (some
               (fn [operator]
                 (when (and
                        (=ci operator-str (:s operator))
                        (= [type sub-type] (:e operator)))
                   (:f operator)))
               OPERATORS-PRECEDENCE)]
     (if o-fn o-fn 'missing-fn))))

(defn is-operator? [test-var]
  (when (map? test-var)
    (let [{:keys [type sub-type value]} test-var]
      (some #(=ci (:s %) value) OPERATORS-PRECEDENCE))))

(defn get-higher-precendence [{type-1 :type sub-type-1 :sub-type value-1 :value :as o1}
                              {type-2 :type sub-type-2 :sub-type value-2 :value :as o2}]
  (let [c-1 (case type-1
              :OperatorInfix :infix
              :OperatorPrefix :prefix
              :OperatorPostfix :postfix
              :Function :args
              nil)
        c-2 (case type-2
              :OperatorInfix :infix
              :OperatorPrefix :prefix
              :OperatorPostfix :postfix
              :Function :args
              nil)]
    (some #(cond
             (and
              (=ci (:s %) value-1)
              (= (:c %) c-1)) o1
             (and
              (=ci (:s %) value-2)
              (= (:c %) c-2)) o2
             :else nil)
          OPERATORS-PRECEDENCE)))

(comment
  (is-operator? "+")
  (get-higher-precendence "*" "+")
  (get-higher-precendence {:value "+" :type :OperatorInfix} {:value "-" :type :OperatorPrefix})
  (get-higher-precendence {:value "*" :type :OperatorInfix} {:value "+" :type :OperatorInfix})
  (get-higher-precendence {:value "%" :type :OperatorPostfix} {:value "*" :type :OperatorInfix})
  (split-at 3 (list 1 2 3)))

(defn pop-to-expression [operator-stack operand-stack & [final?]]
  (loop [operator-stack operator-stack
         operand-stack operand-stack
         idx 0]
    (println-dbg "POP/ OAND" operand-stack)
    (println-dbg "POP/ OSTK" operator-stack)
    (if (or (not (seq operator-stack))
            (> idx 10))
      (do
        (println-dbg "POP/ DONE")
        (println-dbg "")
        [operand-stack operator-stack])
      (let [{:keys [type sub-type value] :as operator}
            (first operator-stack)
            args-count-param (get-args-count type sub-type value)
            args-count (if (= :all args-count-param) (count operand-stack) args-count-param)
            [args rst] (split-at
                        args-count
                        operand-stack)]
        (println-dbg "POP/ OP" operator)
        (println-dbg "POP/ ACOUNT" args-count)
        (println-dbg "POP/ ARGS" args)
        (println-dbg "POP/ RST" rst)
        (println-dbg "POP/ ADD" (list (conj (reverse args) operator)))
        (println-dbg "")

        (recur
         (rest operator-stack)
         (concat #_rst
                 (list (conj (reverse args) operator))
                 rst)
         (if final? (inc idx) 2000))))))

(comment
  {:vlaaad.reveal/command '(clear-output)}
  :end)

(defn parse-simple-expression-tokens [tokens]
  (loop [input tokens operator-stack (list) operand-stack (list)]
    (if-not (seq input)
      (ffirst (pop-to-expression operator-stack operand-stack true))
      (let [term (first input)
            operator? (is-operator? term)
            action (cond (not operator?)
                         :push-operand
                         (= term (get-higher-precendence (peek operator-stack) term))
                         :push-operator
                         :else
                         :reset-stack)]
        (println-dbg "INPUT" input)
        (println-dbg "OAND" operand-stack)
        (println-dbg "OSTK" operator-stack)
        (println-dbg "TERM" term)
        (println-dbg "ACTION" action)
        (println-dbg "")

        (when (= :reset-stack action)
          (println-dbg "RESET OAND" operand-stack)
          (println-dbg "RESET OSTK" operator-stack)
          (println-dbg "RESET TERM" term)
          (println-dbg "RESET COMP" (peek operator-stack))
          (println-dbg "RESET TREE" (list (concat operator-stack operand-stack)))
          (println-dbg ""))

        (if (= :reset-stack action)
          (let [[new-operand-stack new-operator-stack]
                (pop-to-expression operator-stack operand-stack)]
            (recur
             input
             new-operator-stack
             new-operand-stack))
          (recur
           (rest input)
           (case action
             :push-operator (conj operator-stack term)
             operator-stack)
           (case action
             :push-operand (conj operand-stack term)
             operand-stack)))))))

(defn parse-expression-tokens
  [tokens]
  (walk/postwalk
   (fn [form]
     (if (and (vector? form)
              (not (map-entry? form)))
       (parse-simple-expression-tokens form)
       form))
   tokens))

(comment
  (parse-expression-tokens [{:sub-type nil
                             :type :OperatorPrefix
                             :value "-"}
                            {:sub-type :Number
                             :type :Operand
                             :value "2"}
                            {:sub-type :Math
                             :type :OperatorInfix
                             :value "+"}
                            {:sub-type :Number
                             :type :Operand
                             :value "2"}])

  (parse-simple-expression-tokens [{:sub-type nil
                                    :type :OperatorPrefix
                                    :value "-"}
                                   {:sub-type :Number
                                    :type :Operand
                                    :value "2"}
                                   {:sub-type :Math
                                    :type :OperatorInfix
                                    :value "+"}
                                   {:sub-type :Number
                                    :type :Operand
                                    :value "2"}])

  (parse-simple-expression-tokens [{:value "max"
                                    :type :Function
                                    :sub-type :Start}
                                   {:sub-type :Number
                                    :type :Operand
                                    :value "1"}
                                   {:sub-type :Number
                                    :type :Operand
                                    :value "2"}])

  (parse-expression-tokens [{:value "max"
                             :type :Function
                             :sub-type :Start}
                            {:sub-type :Number
                             :type :Operand
                             :value "1"}
                            [{:sub-type nil
                              :type :OperatorPrefix
                              :value "-"}
                             {:sub-type :Number
                              :type :Operand
                              :value "2"}
                             {:sub-type :Math
                              :type :OperatorInfix
                              :value "+"}
                             {:sub-type :Number
                              :type :Operand
                              :value "3"}]])

  (parse-expression-tokens [{:value "max"
                             :type :Function
                             :sub-type :Start}
                            {:sub-type :Number
                             :type :Operand
                             :value "1"}
                            [[{:sub-type nil
                               :type :OperatorPrefix
                               :value "-"}
                              {:sub-type :Number
                               :type :Operand
                               :value "2"}]
                             {:sub-type :Math
                              :type :OperatorInfix
                              :value "+"}
                             {:sub-type :Number
                              :type :Operand
                              :value "3"}]])

  (is-operator? {:sub-type nil
                 :type :OperatorPrefix
                 :value "-"})

  :end)