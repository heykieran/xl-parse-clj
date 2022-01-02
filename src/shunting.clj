(ns shunting
  (:require [clojure.string :as str]
            [clojure.walk :as walk]))

(def OPERATORS-PRECEDENCE
  [{:name :unary-plus :s "+" :f 'functions/fn-unary-plus :c :prefix :a 1 :e [:OperatorPrefix nil]}
   {:name :unary-minus :s "-" :f 'functions/fn-unary-minus :c :prefix :a 1 :e [:OperatorPrefix nil]}

   {:name :unary-prcnt :s "%" :f 'functions/prcnt :c :postfix :a 1 :e [:OperatorPostfix nil]}

   {:name :union :s "," :f 'functions/fn-union :c :infix :a :all :e [:OperatorInfix :Union]}

   {:name :binary-mult :s "*" :f 'functions/fn-multiply :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-div :s "/" :f 'functions/fn-divide  :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-plus :s "+" :f 'functions/fn-add :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-minus :s "-" :f 'functions/fn-subtract :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-exp :s "^" :f 'functions/fn-exponent :c :infix :a 2 :e [:OperatorInfix :Math]}
   {:name :binary-concat :s "&" :f 'functions/fn-concat :c :infix :a 2 :e [:OperatorInfix :Concatenation]} ; what should this do? There are coercions to text in Excel

   {:name :compare-eq :s "=" :f 'functions/fn-equal? :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-gt :s ">" :f 'functions/fn-gt? :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-lt :s "<" :f 'functions/fn-lt? :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-gte :s ">=" :f 'functions/fn-gt-equal? :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-lte :s "<=" :f 'functions/fn-lt-equal? :c :infix :a 2 :e [:OperatorInfix :Logical]}
   {:name :compare-neq :s "<>" :f 'functions/fn-not-equal? :c :infix :a 2 :e [:OperatorInfix :Logical]}

   {:name :index :s "index" :f 'functions/fn-index :c :args :a :all :e [:Function :Start]}

   {:name :abs :s "abs" :f 'functions/abs :c :args :a 1 :e [:Function :Start]}
   {:name :sin :s "sin" :f 'Math/sin :c :args :a 1 :e [:Function :Start]}
   {:name :true :s "true" :f 'functions/fn-true :c :args :a 0 :e [:Function :Start]}
   {:name :false :s "false" :f 'functions/fn-false :c :args :a 0 :e [:Function :Start]}
   {:name :and :s "and" :f 'functions/fn-and :c :args :a :all :e [:Function :Start]}
   {:name :or :s "or" :f 'functions/fn-or :c :args :a :all :e [:Function :Start]}
   {:name :not :s "not" :f 'functions/fn-not :c :args :a 1 :e [:Function :Start]}
   {:name :search :s "search" :f 'functions/fn-search :c :args :a 3 :e [:Function :Start]}
   {:name :max :s "max" :f 'functions/fn-max :c :args :a :all :e [:Function :Start]}
   {:name :min :s "min" :f 'functions/fn-min :c :args :a :all :e [:Function :Start]}
   {:name :pi :s "pi" :f 'functions/pi :c :args :a 0 :e [:Function :Start]}
   {:name :sum :s "sum" :f 'functions/fn-sum :c :args :a :all :e [:Function :Start]}
   {:name :sumif :s "sumif" :f 'functions/fn-sumif :c :args :a :all :e [:Function :Start]}
   {:name :average :s "average" :f 'functions/fn-average :c :args :a :all :e [:Function :Start]}
   {:name :averageif :s "averageif" :f 'functions/fn-average-if :c :args :a :all :e [:Function :Start]}
   {:name :count :s "count" :f 'functions/fn-count :c :args :a :all :e [:Function :Start]}
   {:name :countif :s "countif" :f 'functions/fn-count-if :c :args :a :all :e [:Function :Start]}
   {:name :counta :s "counta" :f 'functions/fn-counta :c :args :a :all :e [:Function :Start]}
   {:name :concatenate :s "concatenate" :f 'functions/fn-concatenate :c :args :a :all :e [:Function :Start]}
   {:name :now :s "now" :f 'functions/fn-now :c :args :a 0 :e [:Function :Start]}
   {:name :date :s "date" :f 'functions/fn-date :c :args :a :all :e [:Function :Start]}
   {:name :days :ext true :s "_xlfn.days" :f 'functions/fn-days :c :args :a :all :e [:Function :Start]}
   {:name :datevalue :s "datevalue" :f 'functions/fn-datevalue :c :args :a 1 :e [:Function :Start]}
   {:name :yearfrac :s "yearfrac" :f 'functions/fn-yearfrac :c :args :a :all :e [:Function :Start]}
   {:name :match :s "match" :f 'functions/fn-match :c :args :a :all :e [:Function :Start]}
   {:name :indirect :s "indirect" :f 'functions/fn-indirect :c :args :a :all :e [:Function :Start] :context-arg true}
   {:name :offset :s "offset" :f 'functions/fn-offset :c :args :a :all :e [:Function :Start] :context-arg true}
   {:name :vlookup :s "vlookup" :f 'functions/fn-vlookup :c :args :a :all :e [:Function :Start]}

   {:name :if :s "if" :f 'if :c :args :a 3 :e [:Function :Start]}

   {:name :binary-colon :s ":" :f 'functions/fn-range :c :infix :a 2 :e [:OperatorInfix :Math]}])

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
  ([operator-str type sub-type sheet-name]
   (let [o-fn (some
               (fn [operator]
                 (when (and
                        (=ci operator-str (:s operator))
                        (= [type sub-type] (:e operator)))
                   operator))
               OPERATORS-PRECEDENCE)]
     (if o-fn
       (if (:context-arg o-fn)
         ;; if the function takes the sheet as its first argument
         ;; and the context as its second
         (list 'partial (:f o-fn) sheet-name 'graph/*context*)
         (:f o-fn))
       'missing-fn))))

(defn is-operator? [test-var]
  (when (map? test-var)
    (let [{:keys [type sub-type value]} test-var
          is-operator-result (some #(and
                                     (not (= :Operand type))
                                     (=ci (:s %) value))
                                   OPERATORS-PRECEDENCE)]
      (when (and (contains? #{:Function :OperatorInfix :OperatorPostfix :OperatorPrefix} type)
                 (not is-operator-result))
        (throw (IllegalArgumentException. (str "Unknown operator encountered \"" value "\""))))
      is-operator-result)))

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

  (parse-simple-expression-tokens [{:sub-type :Text, :type :Operand, :value ">"}
                                   {:sub-type :Concatenation, :type :OperatorInfix, :value "&"}
                                   {:sub-type :Range, :type :Operand, :value "A1"}])

  (is-operator? {:sub-type :Text, :type :Operand, :value ">"})

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