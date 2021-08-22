(ns shunting-xl
  (:require [clojure.string :as str]
            [clojure.walk :as walk]))

(def OPERATORS-PRECEDENCE
  [{:v "+"   :name :unary-plus :t :OperatorPrefix :f '+ :c 1}
   {:v "-"   :name :unary-minus :t :OperatorPrefix :f '- :c 1}
   {:v "*"   :name :binary-mult :t :OperatorInfix :f '* :c 2}
   {:v "/"   :name :binary-div :t :OperatorInfix :f '/  :c 2}
   {:v "+"   :name :binary-plus :t :OperatorInfix :f '+ :c 2}
   {:v "-"   :name :binary-minus :t :OperatorInfix :f '- :c 2}
   {:v "sin" :name :sin :t :Function :f 'sin :c 1}
   {:v "max" :name :max :t :Function :f 'max :c :all}])

(defn is-operator? [{:keys [type] :as term}]
  (if (contains? #{:OperatorInfix
                   :OperatorPostfix
                   :OperatorPrefix
                   :Function}
                 type)
    true
    false))

(defn get-higher-precendence
  [{o1-type :type o1-value :value :as o1}
   {o2-type :type o2-value :value :as o2}]
  (some (fn[{o-type :t o-value :v}]
          (cond
           (and (= o-type o1-type) (= o-value o1-value)) o1
           (and (= o-type o2-type) (= o-value o2-value)) o2
           :else nil))
        OPERATORS-PRECEDENCE))

(defn get-args-for-operator
  [{:keys [value type sub-type] :as operator} operand-stack]
  (println "GET ARGS" operator)
  (println "STACK" operand-stack)
  (let [[args rst]
        (if (identical? :Function type)
          (split-at (count operand-stack) operand-stack)
          #_(vector
           (keep
            (fn [[k v]]
              (when
               (number? k)
                v))
            operator)
           operand-stack)
          (split-at
           (cond
             (= "max" operator)
             (count operand-stack)
             (identical? :OperatorInfix type)
             2
             :else
             (throw (Exception. "NO ARG PROCESSING AVAILABLE FOR OPERATOR")))
           operand-stack))]
    (println "ARGS" args)
    (println "REST" rst)
    [args rst]))

(defn pop-to-expression [operator-stack operand-stack & [debug]]
  (loop [operator-stack operator-stack
         operand-stack operand-stack]
    (when debug
      (println "POP OAND" operand-stack "OPTR" operator-stack))
    (if-not (seq operator-stack)
      operand-stack
      (let [operator (first operator-stack)
            [args rst] (get-args-for-operator operator operand-stack)]
        (when debug
          (println "OPERATOR" operator)
          (println "ARGS" args))
        (recur
         (rest operator-stack)
         (concat rst (list (conj (reverse args) operator))))))))

(defn parse-simple-expression-tokens [tokens & [debug]]
  (loop [input tokens operator-stack (list) operand-stack (list)]
    (if-not (seq input)
      (first (pop-to-expression operator-stack operand-stack debug))
      (let [term (first input)
            operator? (is-operator? term)
            action (cond (not operator?)
                         :push-operand
                         (= term (get-higher-precendence (peek operator-stack) term))
                         :push-operator
                         :else
                         :reset-stack)]
        (when debug
          (println "INPUT" input "OAND" operand-stack "OPTR" operator-stack "TERM" term))
        (when (and debug (= :reset-stack action))
          (println "RESET OAND" operand-stack "OPTR" operator-stack "TERM" term)
          (println "TREE" (list (concat operator-stack operand-stack))))
        (if (= :reset-stack action)
          (recur
           input
           (pop operator-stack)
           (pop-to-expression operator-stack operand-stack debug))
          (recur
           (rest input)
           (case action
             :push-operator (conj operator-stack term)
             operator-stack)
           (case action
             :push-operand (conj operand-stack term)
             operand-stack)))))))

(defn parse-expression-tokens [tokens & flags]
  (walk/postwalk
   (fn [form]
     (if (vector? form)
       (parse-simple-expression-tokens form flags)
       form))
   tokens))
