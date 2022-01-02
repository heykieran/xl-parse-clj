(ns expressions
  (:require
   [clojure.string :as str]
   [xlparse :as parse]
   [ast-processing :as ast]
   [clojure.walk :as walk]
   [shunting :as sh]))

(defn like-expr->re-str [like-expr]
  (-> like-expr
      (str/replace
       #"((?<!~)\*)"
       ".*")
      (str/replace
       #"(~\*)"
       "\\\\*")
      (str/replace
       #"((?<!~)\?)"
       ".")
      (str/replace
       #"(~\?)"
       "\\\\?")))

(defn recast-comparative-expression [text-expr]
  (let [[has-comparative? comp-op] (re-matches #"^(<(?!>)|(?<!<)>|=|<>).*$" (str text-expr))
        is-str? (string? text-expr)
        wildcard-exp? (and is-str?
                           (re-matches #".*((?<!~)\*|(?<!~)\?).*" text-expr))]
    (if has-comparative?
      (str "$SELF" comp-op (subs text-expr (count comp-op)))
      (str "$SELF" (if wildcard-exp? " = " "=")
           (if is-str?
             (if wildcard-exp?
               (pr-str (-> text-expr like-expr->re-str (str "__as_regex")))
               (-> text-expr
                   (str/replace
                    #"(~\*)"
                    "*")
                   (str/replace
                    #"(~\?)"
                    "?")
                   (pr-str)))
             text-expr)))))

(defn ->code [expr]
  (-> (str "=" expr)
      (parse/parse-to-tokens)
      (parse/nest-ast)
      (parse/wrap-ast)
      (ast/process-tree)
      (sh/parse-expression-tokens)
      (ast/unroll-for-code-form)))

(defn code->with-regex [form]
  (walk/postwalk
   (fn [t]
     (if (and (list? t)
              (= 2 (count t))
              (= 'str (first t))
              (string? (last t))
              (re-matches #"^.*__as_regex$" (last t)))
       `(re-pattern (str/replace ~(last t) #"__as_regex$" ""))
       t))
   form))

(defn reduce-by-comp-expression [expr-form search-seq & [val-seq]]
  (let [val-seq (or val-seq
                    search-seq)]
    (loop [s-seq (some-> search-seq deref) v-seq (some-> val-seq deref) filtered-seq []]
      (if-not (seq s-seq)
        filtered-seq
        (let [s-val (first s-seq)
              v-val (first v-seq)
              match? (eval (walk/postwalk-replace
                            {'(eval-range
                               "$SELF") s-val}
                            expr-form))]
          (recur
           (rest s-seq)
           (rest v-seq)
           (if match? (conj filtered-seq v-val) filtered-seq)))))))

(comment
  (re-matches #"^(<(?!>)|(?<!<)>|=|<>).*$" ">100.0")
  (re-matches #"^(<(?!>)|(?<!<)>|=|<>).*$" "2*")

  (recast-comparative-expression "2*")
  (recast-comparative-expression "2~*")
  (recast-comparative-expression "2~**")
  (recast-comparative-expression ">100.0")
  (recast-comparative-expression "100.0")
  (recast-comparative-expression 100.0)

  (-> "2*"
      (recast-comparative-expression)
      (->code)
      (code->with-regex)))