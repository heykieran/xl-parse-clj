(ns xlparse
  (:require [clojure.string :as str]
            [clojure.walk :as walk]))

(def QUOTE-DOUBLE \")
(def QUOTE-SINGLE \')
(def BRACKET-CLOSE \])
(def BRACKET-OPEN \[)
(def BRACE-OPEN \{)
(def BRACE-CLOSE \})
(def PAREN-OPEN \()
(def PAREN-CLOSE \))
(def SEMICOLON \;)
(def WHITESPACE \space)
(def COMMA \,)
(def ERROR-START \#)

(def OPERATORS-SN #{\+ \-})
(def OPERATORS-INFIX #{\+ \- \* \/ \^ \& \= \> \< \:})
(def OPERATORS-POSTFIX #{\%})

(def ERRORS #{"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"})

(def COMPARATORS-MULTI [">=", "<=", "<>"])

(defn convert-double [number-string]
  (try
    (Double/parseDouble number-string)
    (catch Exception _)))

(defn peek-stop-token-from-stack [stack]
  (when-let [v (peek stack)]
    {:value ""
     :type (:type v)
     :sub-type :Stop}))

(defn parse-to-tokens-pass-1 [input-formula]
  (let [formula (cond (str/starts-with? input-formula "=")
                      (subs input-formula 1)
                      (str/starts-with? input-formula "{")
                      input-formula
                      :else
                      (throw (IllegalArgumentException. (str "Input formula must start with '=' or '{'"))))]
    (loop [formula (seq formula)
           tokens []
           stack (list)
           gathered-value ""
           in-string? false
           in-path? false
           in-range? false
           in-error? false]
      (if-not (seq formula)
        (if (> (count gathered-value) 0)
          (conj tokens
                {:value gathered-value
                 :type :Operand})
          tokens)
        (let [chr (first formula)]
          (cond in-string?
                (let [is-quote? (= QUOTE-DOUBLE chr)
                      embedded-quote? (and
                                       (> (count formula) 2)
                                       (= QUOTE-DOUBLE (second formula)))]
                  (recur
                   (if (and is-quote? embedded-quote?)
                     (rest (rest formula))
                     (rest formula))
                   (if (and is-quote? (not embedded-quote?))
                     (conj tokens {:value gathered-value
                                   :type :Operand
                                   :sub-type :Text})
                     tokens)
                   stack
                   (if is-quote?
                     (if embedded-quote?
                       (str gathered-value QUOTE-DOUBLE)
                       "")
                     (str gathered-value chr))
                   (if (and is-quote? (not embedded-quote?))
                     false
                     in-string?)
                   in-path?
                   in-range?
                   in-error?))

                in-path?
                (let [is-single-quote? (= QUOTE-SINGLE chr)
                      embedded-quote? (and
                                       (> (count formula) 2)
                                       (= QUOTE-SINGLE (second formula)))]
                  (recur
                   (if (and is-single-quote? embedded-quote?)
                     (rest (rest formula))
                     (rest formula))
                   tokens
                   stack
                   (if is-single-quote?
                     (if embedded-quote?
                       (str gathered-value QUOTE-SINGLE)
                       gathered-value)
                     (str gathered-value chr))
                   in-string?
                   (if (and is-single-quote? (not embedded-quote?))
                     false
                     in-path?)
                   in-range?
                   in-error?))

                in-range?
                (recur
                 (rest formula)
                 tokens
                 stack
                 (str gathered-value chr)
                 in-string?
                 in-path?
                 (if (= BRACKET-CLOSE chr) false in-range?)
                 in-error?)

                in-error?
                (let [current-error-value (str gathered-value chr)
                      error-token? (contains? ERRORS current-error-value)]
                  (recur
                   (rest formula)
                   (cond->
                    tokens
                     error-token?
                     (conj
                      {:value current-error-value
                       :type :Operand
                       :sub-type :Error}))
                   stack
                   (if error-token? "" current-error-value)
                   in-string?
                   in-path?
                   in-range?
                   (if error-token? false in-error?)))

                (and
                 (contains? OPERATORS-SN chr)
                 (> (count gathered-value) 0)
                 (re-matches #"^[1-9]{1}(\.[0-9]+)?E{1}$" gathered-value))
                (recur
                 (rest formula)
                 tokens
                 stack
                 (str gathered-value chr)
                 in-string?
                 in-path?
                 in-range?
                 in-error?)

                (= QUOTE-DOUBLE chr)
                (recur
                 (rest formula)
                 (cond-> tokens
                   (> (count gathered-value) 0)
                   (conj
                    {:value ""
                     :type :Unknown}))
                 stack
                 gathered-value
                 true ; in-string
                 in-path?
                 in-range?
                 in-error?)

                (= QUOTE-SINGLE chr)
                (recur
                 (rest formula)
                 (cond-> tokens
                   (> (count gathered-value) 0)
                   (conj
                    {:value ""
                     :type :Unknown}))
                 stack
                 gathered-value
                 in-string?
                 true ; in-path
                 in-range?
                 in-error?)

                (= BRACKET-OPEN chr)
                (recur
                 (rest formula)
                 tokens
                 stack
                 (str gathered-value BRACKET-OPEN)
                 in-string?
                 in-path?
                 true
                 in-error?)

                (= ERROR-START chr)
                (recur
                 (rest formula)
                 (cond-> tokens
                   (> (count gathered-value) 0)
                   (conj
                    {:value ""
                     :type :Unknown}))
                 stack
                 (if (> (count gathered-value) 0)
                   ""
                   (str gathered-value ERROR-START))
                 in-string?
                 in-path?
                 in-range?
                 true)

                (= BRACE-OPEN chr)
                (let
                 [value? (> (count gathered-value) 0)
                  array-token {:value "ARRAY"
                               :type :Function
                               :sub-type :Start}
                  array-row-token {:value "ARRAYROW"
                                   :type :Function
                                   :sub-type :Start}]
                  (recur
                   (rest formula)
                   (cond->
                    tokens
                     value?
                     (conj
                      {:value gathered-value
                       :type :Unknown})
                     true
                     (conj
                      array-token)
                     true
                     (conj
                      array-row-token))
                   (conj stack array-token array-row-token)
                   (if value? "" gathered-value)
                   in-string?
                   in-path?
                   in-range?
                   in-error?))

                (= SEMICOLON chr)
                (let [value? (> (count gathered-value) 0)
                      array-row-token {:value "ARRAYROW"
                                       :type :Function
                                       :sub-type :Start}]
                  (recur
                   (rest formula)
                   (cond-> tokens
                     value?
                     (conj
                      {:value gathered-value
                       :type :Operand})
                     true
                     (conj
                      (peek-stop-token-from-stack stack))
                     true
                     (conj
                      {:value ","
                       :type :Argument})
                     true
                     (conj
                      array-row-token))
                   (conj
                    (pop stack)
                    array-row-token)
                   (if value? "" gathered-value)
                   in-string?
                   in-path?
                   in-range?
                   in-error?))

                (= BRACE-CLOSE chr)
                (let
                 [value? (> (count gathered-value) 0)
                  stop-token-1 (peek-stop-token-from-stack stack)
                  stop-token-2 (peek-stop-token-from-stack (pop stack))]
                  (recur
                   (rest formula)
                   (cond-> tokens
                     value?
                     (conj
                      {:value gathered-value
                       :type :Operand})
                     true
                     (conj
                      stop-token-1)
                     true
                     (conj
                      stop-token-2))
                   (pop stack)
                   (if value? "" gathered-value)
                   in-string?
                   in-path?
                   in-range?
                   in-error?))

                (= WHITESPACE chr)
                (recur
                 (drop-while #(= \space %) formula)
                 (cond-> tokens
                   (> (count gathered-value) 0)
                   (conj
                    {:value gathered-value
                     :type :Operand})
                   true
                   (conj
                    {:value ""
                     :type :Whitespace}))
                 stack
                 (if (> (count gathered-value) 0) "" gathered-value)
                 in-string?
                 in-path?
                 in-range?
                 in-error?)

                (and
                 (>= (count formula) 2)
                 (some #(= % (str chr (second formula))) COMPARATORS-MULTI))
                (recur
                 (rest (rest formula))
                 (cond->
                  tokens
                   (> (count gathered-value) 0)
                   (conj
                    {:value gathered-value
                     :type :Operand})
                   true
                   (conj
                    {:value (str chr (second formula))
                     :type :OperatorInfix
                     :sub-type :Logical}))
                 stack
                 (if (> (count gathered-value) 0) "" gathered-value)
                 in-string?
                 in-path?
                 in-range?
                 in-error?)

                (contains? OPERATORS-INFIX chr)
                (recur
                 (rest formula)
                 (cond-> tokens
                   (> (count gathered-value) 0)
                   (conj
                    {:value gathered-value
                     :type :Operand})
                   true
                   (conj
                    {:value (str chr)
                     :type :OperatorInfix}))
                 stack
                 (if (> (count gathered-value) 0) "" gathered-value)
                 in-string?
                 in-path?
                 in-range?
                 in-error?)

                (contains? OPERATORS-POSTFIX chr)
                (let [value? (> (count gathered-value) 0)]
                  (recur
                   (rest formula)
                   (cond-> tokens
                     value?
                     (conj
                      {:value gathered-value
                       :type :Operand})
                     true
                     (conj
                      {:value (str chr)
                       :type :OperatorPostfix}))
                   stack
                   (if value? "" gathered-value)
                   in-string?
                   in-path?
                   in-range?
                   in-error?))

                (= PAREN-OPEN chr)
                (let [value? (> (count gathered-value) 0)
                      start-token (if value?
                                    {:value gathered-value
                                     :type :Function
                                     :sub-type :Start}
                                    {:value ""
                                     :type :Subexpression
                                     :sub-type :Start})]
                  (recur
                   (rest formula)
                   (conj
                    tokens
                    start-token)
                   (conj stack start-token)
                   (if value? "" gathered-value)
                   in-string?
                   in-path?
                   in-range?
                   in-error?))

                (= COMMA chr)
                (recur
                 (rest formula)
                 (cond-> tokens
                   (> (count gathered-value) 0)
                   (conj
                    {:value gathered-value
                     :type :Operand})
                   (not= :Function (-> stack peek :type))
                   (conj
                    {:value ","
                     :type :OperatorInfix
                     :sub-type :Union})
                   (= :Function (-> stack peek :type))
                   (conj
                    {:value ","
                     :type :Argument}))
                 stack
                 (if (> (count gathered-value) 0) "" gathered-value)
                 in-string?
                 in-path?
                 in-range?
                 in-error?)

                (= PAREN-CLOSE chr)
                (recur
                 (rest formula)
                 (cond-> tokens
                   (> (count gathered-value) 0)
                   (conj
                    {:value gathered-value
                     :type :Operand})
                   true
                   (conj
                    (peek-stop-token-from-stack stack)))
                 (pop stack)
                 (if (> (count gathered-value) 0) "" gathered-value)
                 in-string?
                 in-path?
                 in-range?
                 in-error?)

                :else

                (recur
                 (rest formula)
                 tokens
                 stack
                 (str gathered-value chr)
                 in-string?
                 in-path?
                 in-range?
                 in-error?)))))))

(defn parse-to-tokens-pass-2 [tokens]
  (reduce
   (fn [accum [{prev-type :type prev-sub-type :sub-type :as prev-token}
               {current-type :type :as current-token}
               {next-type :type next-sub-type :sub-type :as next-token}]]
     (cond (nil? current-token)
           accum

           (not= :Whitespace current-type)
           (conj accum current-token)

           (nil? prev-token)
           accum

           (not
            (or (and (= :Function prev-type) (= :Stop prev-sub-type))
                (and (= :Subexpression prev-type) (= :Stop prev-sub-type))
                (= :Operand prev-type)))
           accum

           (nil? next-token)
           accum

           (not
            (or (and (= :Function next-type) (= :Stop next-sub-type))
                (and (= :Subexpression next-type) (= :Stop next-sub-type))
                (= :Operand next-type)))
           accum

           :else
           (conj accum {:value ""
                            :type :OperatorInfix
                            :sub-type :Intersection})))
   []
   (partition 3 1 (into [] (concat [nil] tokens [nil])))))

(defn parse-to-tokens-pass-3 [tokens]
  (reduce
   (fn [accum [{prev-type :type prev-sub-type :sub-type :as prev-token}
               {current-type :type current-sub-type :sub-type current-value :value :as current-token}
               _]]
     (cond
       (nil? current-token)
       accum

       (and (= :OperatorInfix current-type) (= "-" current-value))
       (let [first? (nil? prev-token)
             prev-continue?
             (or
              (and (= :Function prev-type) (= :Stop prev-sub-type))
              (and (= :Subexpression prev-type) (= :Stop prev-sub-type))
              (= :OperatorPostfix prev-type)
              (= :Operand prev-type))]
         (conj accum
                   {:value current-value
                    :type (cond
                            first?
                            :OperatorPrefix
                            prev-continue?
                            current-type
                            :else
                            :OperatorPrefix)
                    :sub-type (if prev-continue? :Math current-sub-type)}))

       (and (= :OperatorInfix current-type) (= "+" current-value))
       (let [first? (nil? prev-token)
             prev-continue?
             (or
              (and (= :Function prev-type) (= :Stop prev-sub-type))
              (and (= :Subexpression prev-type) (= :Stop prev-sub-type))
              (= :OperatorPostfix prev-type)
              (= :Operand prev-type))]
         (conj accum
                   {:value current-value
                    :type  current-type
                    :sub-type (cond
                                first?
                                current-sub-type
                                prev-continue?
                                :Math
                                :else
                                current-sub-type)}))

       (and (= :OperatorInfix current-type) (nil? current-sub-type))
       (conj accum
                 {:value current-value
                  :type  current-type
                  :sub-type (cond
                              (str/includes? "<>=" (str (first current-value)))
                              :Logical
                              (= "&" current-value)
                              :Concatenation
                              :else
                              :Math)})

       (and (= :Operand current-type) (nil? current-sub-type))
       (let [double-number (convert-double current-value)]
         (conj accum
                   {:value current-value
                    :type  current-type
                    :sub-type (cond
                                (some? double-number)
                                :Number
                                (contains? #{"TRUE" "FALSE"} current-value)
                                :Logical
                                :else
                                :Range)}))

       :else
       (conj
        accum
        {:value (if
                 (and
                  (= :Function current-type)
                  (> (count current-value) 0)
                  (str/starts-with? current-value "@"))
                  (subs current-value 1)
                  current-value)
         :type  current-type
         :sub-type current-sub-type})))
   []
   (partition 3 1 (into [] (concat [nil] tokens [nil])))))

(defn parse-to-tokens-pass-4 [tokens]
  (loop [part-tokens (partition 3 1 (into [] (concat [nil] tokens [nil]))) result-vec []]
    (if-not (seq part-tokens)
      result-vec
      (let [[{prev-type :type prev-sub-type :sub-type prev-value :value :as prev-token}
             {current-type :type current-sub-type :sub-type current-value :value :as current-token}
             {next-type :type next-sub-type :sub-type next-value :value :as next-token}]
            (first part-tokens)
            ;; because we treat ':' as an infix operator to support INDEX and OFFEST etc.
            ;; we might need to reassemble a 'real' range that was earlier decomposed.
            consolidate? (and (contains? #{:Range :Number} prev-sub-type)
                              (= current-type :OperatorInfix)
                              (= ":" current-value)
                              (contains? #{:Range :Number} next-sub-type))]
        (recur (cond-> part-tokens
                 true
                 rest
                 consolidate?
                 rest)
               (cond
                 (nil? current-token)
                 result-vec
                 consolidate?
                 (conj
                  (into [] (butlast result-vec))
                  {:type :Operand
                   :sub-type :Range
                   :value (str prev-value current-value next-value)})
                 :else
                 (conj result-vec current-token)))))))

(defn parse-to-tokens [formula]
  (-> formula
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2)
      (parse-to-tokens-pass-3)
      (parse-to-tokens-pass-4)))

(defn map-comparator [v1 v2]
  (cond
    (identical? (type v1) (type v2))
    (compare v1 v2)
    (and (keyword? v1) (not (keyword? v2)))
    -1
    (and (not (keyword? v1)) (keyword? v2))
    1))

(defn sort-nested-ast
  [nested-ast]
  (walk/postwalk
   (fn [item]
     (if (map? item)
       (into (sorted-map-by map-comparator) item)
       item))
   nested-ast))

(defn nest-ast [ast]
  (loop [ast ast path [0] idx 1 result {}]
    (if-not (seq ast)
      (sort-nested-ast result)
      (let [{node-sub-type :sub-type :as node}
            (first ast)
            node-id idx
            is-start? (= :Start node-sub-type)
            is-end? (= :Stop node-sub-type)
            new-path (cond
                       is-start? (conj path node-id)
                       is-end? (into [] (butlast path))
                       :else path)]
        (recur
         (rest ast)
         new-path
         (inc idx)
         (if-not is-end?
           (-> result
               (assoc-in
                (conj path node-id)
                node))
           result))))))

(defn wrap-ast
  "Make the top-level a subexpression
   if it contains more than a single entry"
  [ast]
  (if (not= 1 (count (vals (get ast 0))))
    (update ast 0
            assoc
            :value ""
            :type :Subexpression
            :sub-type :Start)
    ast))

(comment

  (convert-double "113.222A")
  (parse-to-tokens-pass-1 "= IF(1>1, \"YES\", \"NO\")")
  (parse-to-tokens-pass-1 "= 1 + $A$1 + 33 +  \"HELLO\"")
  (parse-to-tokens-pass-1 "= IF(1=#NUM!, \"YES\", \"NO\")")
  (parse-to-tokens-pass-1 "=SUM(C5:C8)")
  (parse-to-tokens-pass-1 "=20% * 100")
  (parse-to-tokens-pass-1 "=20% * 1E2")

  (-> "= IF(1=1, \"YES\", \"NO\")"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2))

  (partition 3 1 (into [] (concat [nil]
                                  (-> "= IF(1=1, \"YES\", \"NO\")"
                                      (parse-to-tokens-pass-1)
                                      #_(parse-to-tokens-pass-2))
                                  [nil])))

  (-> "= 1 + $A$1 + 33 +  \"HELLO\""
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2))

  (-> "= IF(1=#NUM!, \"YES\", \"NO\")"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2))

  (-> "=SUM(C5:C8)"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2))

  (-> "=20% * 100"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2))

  (-> "= IF(1=1, \"YES\", \"NO\")"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2)
      (parse-to-tokens-pass-3))

  (-> "= 1 + $A$1 + 33 +  \"HELLO\""
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2)
      (parse-to-tokens-pass-3))

  (-> "= IF(1=#NUM!, \"YES\", \"NO\")"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2)
      (parse-to-tokens-pass-3))

  (-> "=SUM(C5:C8)"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2)
      (parse-to-tokens-pass-3))

  (-> "=20% * -100"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2)
      (parse-to-tokens-pass-3))

  (-> "=20% * 1E2"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2)
      (parse-to-tokens-pass-3))

  ;;

  (-> "{=MAX(IF(ISERROR(SEARCH(H5&\"*\",files)),0,ROW(files)-ROW(INDEX(files,1,1))+1))}"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2)
      (parse-to-tokens-pass-3))

  (-> "=MAX(1,2+3)"
      (parse-to-tokens-pass-1)
      (parse-to-tokens-pass-2)
      (parse-to-tokens-pass-3))

  (parse-to-tokens "{=MAX(IF(ISERROR(SEARCH(H5&\"*\",files)),0,ROW(files)-ROW(INDEX(files,1,1))+1))}")
  (parse-to-tokens "=(1+2+3)*(10+20+30)/(40+50+60)")

  :end)

(comment
  (def ast-1 [{:value "ARRAY", :type :Function, :sub-type :Start}
            {:value "ARRAYROW", :type :Function, :sub-type :Start}
            {:value "=", :type :OperatorInfix, :sub-type :Logical}
            {:value "MAX", :type :Function, :sub-type :Start}
            {:value "IF", :type :Function, :sub-type :Start}
            {:value "ISERROR", :type :Function, :sub-type :Start}
            {:value "SEARCH", :type :Function, :sub-type :Start}
            {:value "H5", :type :Operand, :sub-type :Range}
            {:value "&", :type :OperatorInfix, :sub-type :Concatenation}
            {:value "*", :type :Operand, :sub-type :Text}
            {:value ",", :type :Argument, :sub-type nil}
            {:value "files", :type :Operand, :sub-type :Range}
            {:value "", :type :Function, :sub-type :Stop}
            {:value "", :type :Function, :sub-type :Stop}
            {:value ",", :type :Argument, :sub-type nil}
            {:value "0", :type :Operand, :sub-type :Number}
            {:value ",", :type :Argument, :sub-type nil}
            {:value "ROW", :type :Function, :sub-type :Start}
            {:value "files", :type :Operand, :sub-type :Range}
            {:value "", :type :Function, :sub-type :Stop}
            {:value "-", :type :OperatorInfix, :sub-type :Math}
            {:value "ROW", :type :Function, :sub-type :Start}
            {:value "INDEX", :type :Function, :sub-type :Start}
            {:value "files", :type :Operand, :sub-type :Range}
            {:value ",", :type :Argument, :sub-type nil}
            {:value "1", :type :Operand, :sub-type :Number}
            {:value ",", :type :Argument, :sub-type nil}
            {:value "1", :type :Operand, :sub-type :Number}
            {:value "", :type :Function, :sub-type :Stop}
            {:value "", :type :Function, :sub-type :Stop}
            {:value "+", :type :OperatorInfix, :sub-type :Math}
            {:value "1", :type :Operand, :sub-type :Number}
            {:value "", :type :Function, :sub-type :Stop}
            {:value "", :type :Function, :sub-type :Stop}
            {:value "", :type :Function, :sub-type :Stop}
            {:value "", :type :Function, :sub-type :Stop}])

  (def ast-2
    [{:value "", :type :Subexpression, :sub-type :Start}
     {:value "1", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "2", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "3", :type :Operand, :sub-type :Number}
     {:value "", :type :Subexpression, :sub-type :Stop}
     {:value "*", :type :OperatorInfix, :sub-type :Math}
     {:value "", :type :Subexpression, :sub-type :Start}
     {:value "10", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "20", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "30", :type :Operand, :sub-type :Number}
     {:value "", :type :Subexpression, :sub-type :Stop}
     {:value "/", :type :OperatorInfix, :sub-type :Math}
     {:value "", :type :Subexpression, :sub-type :Start}
     {:value "40", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "50", :type :Operand, :sub-type :Number}
     {:value "+", :type :OperatorInfix, :sub-type :Math}
     {:value "60", :type :Operand, :sub-type :Number}
     {:value "", :type :Subexpression, :sub-type :Stop}])

  (nest-ast ast-2)

  (tree-seq
   seq?
   identity
   (nest-ast ast-2))

  (walk/postwalk
   (fn [form]
     (if (map? form)
       (do
        (println form)
         form)
       form))
   (nest-ast ast-2))

  (walk/postwalk-demo
   (nest-ast ast-2))

  :end)