(ns ast-processing
  (:require
   [clojure.walk :as walk]
   [shunting :as sh]
   [xlparse :as parse]))

(defn process-function
  "Given a map element representing a function call element in the ast,
   i.e. where :type is :Function, unroll the arguments so they can be
   processed. For example 'max(1,2+3)' would be resolved into the ast
   as a single map, but the element '2+3' needs to be wrapped so it
   can be calculated."
  [fmap]
  (loop [p (into (sorted-map-by parse/map-comparator) fmap) ; don't forget to sort!
         result [(select-keys fmap [:value :type :sub-type]) []]]
    (if-not (seq p)
      result
      (let [[k {:keys [type sub-type value] :as v}] (first p)]
        (recur (rest p)
               (if (number? k)
                 (if (and (= :Argument type) (= "," value))
                   (conj result [])
                   (conj
                    (if (seq result)
                      (pop result)
                      [])
                    (conj
                     (into [] (peek result))
                     v)))
                 result))))))

(defn process-tree
  "Process an AST tree. Currently only treats functions specially to ensure
   that expressions in the function's arguments are processed correctly."
  [ast-tree]
  (walk/postwalk
   (fn [form]
     (if (and (map? form)
              (or (= :Function (:type form))
                  (= :Subexpression (:type form))))
       (process-function form)
       form))
   ast-tree))

(defn process-form-map-to-pseudo-code
  "Given something that looks like a excel parsed element i.e.
   a map containing a :type :sub-type and :value, return
   a reasonable clojure representation"
  ([fmap]
   (process-form-map-to-pseudo-code fmap nil))
  ([{:keys [type sub-type value] :as fmap} sheet-name]
   (cond
     (= [:Operand :Number] [type sub-type])
     (Double/parseDouble value)
     (= [:Operand :Text] [type sub-type])
     (list 'str value)
     (= [:Operand :Range] [type sub-type])
     (list 'eval-range (if sheet-name (str sheet-name "!" value) value))
     (contains? #{:Function :OperatorInfix :OperatorPostfix :OperatorPrefix} type)
     (sh/get-operator-fn value type sub-type sheet-name)
     :else
     fmap)))

(defn unroll-for-code-form
  ([ast-tree]
   (unroll-for-code-form ast-tree nil))
  ([ast-tree sheet-name]
   (let [ast-map (walk/postwalk
                  (fn [form]
                    (if (and (map? form)
                             (every? (fn [k] (some #(= k %) (keys form))) [:type :sub-type :value]))
                      (process-form-map-to-pseudo-code form sheet-name)
                      form))
                  ast-tree)]
     (loop [ast-items ast-map]
       (if (not (map? ast-items))
         ast-items
         (recur (first (vals ast-items))))))))