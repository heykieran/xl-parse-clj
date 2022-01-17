;; Taken with gratitude from StackOverflow answer 
;; https://stackoverflow.com/a/20724760/3864577
;; by jbm

(ns box
  (:import [java.io Writer]))

(deftype Box [value _meta]
  clojure.lang.IObj
  (meta [_] _meta)
  (withMeta [_ m] (Box. value m))
  clojure.lang.IDeref
  (deref [_] value)
  Object
  (toString [this]
    (str (.getName (class this))
         ": "
         (pr-str value))))

(defmethod print-method Box [o, ^Writer w]
  (.write w "#<")
  (.write w (.getName (class o)))
  (.write w ": ")
  (.write w (-> o deref pr-str))
  (.write w ">"))

(defn box
  ([value] (box value nil))
  ([value meta] (Box. value meta)))

(comment
  (instance? Box (box 1 {:a 2}))
  (instance? Box 1))