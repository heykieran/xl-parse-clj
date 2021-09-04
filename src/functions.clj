(ns functions
  (:require
   [clojure.string :as str]
   [excel :as excel]))

(defn abs [v]
  (if (neg? v)
    (- v)
    v))

(defn- v->boolean [v]
  (cond
    (nil? v)
    false
    (and (number? v) (zero? v))
    false
    :else
    (boolean v)))

(defn fn-true []
  true)

(defn fn-false []
  false)

(defn fn-and [& vs]
  (every? #(true? (v->boolean %)) vs))

(defn fn-or [& vs]
  (or (some #(true? (v->boolean %)) vs) false))

(defn fn-not [v]
  (not (v->boolean v)))

(defn prcnt [v]
  (/ (bigdec v) 100.0M))

(defn pi []
  (Math/PI))

(defn sum [& vs]
  (apply + (flatten vs)))

(defn fn-count [& vs]
  (count (keep #(when (number? %) %) vs)))

(defn fn-counta [& vs]
  (-> (keep #(when (not (str/blank? (str %))) %) (flatten vs))
      (count)
      (float)))

(defn average [& vs]
  (/ (apply sum vs)
     (apply fn-count vs)))

(defn fn-now []
  (excel/excel-now))

(defn fn-days [& [d1 d2]]
  (- d1 d2))

(defn fn-yearfrac [& [d1 d2 b]]
  (case b
    (nil 0.) (excel/nasd-360-diff d1 d2)
    1. (/ (- d1 d2) 365.)
    2. (/ (- d1 d2) 365.)
    3. (/ (- d1 d2) 365.)
    4. (/ (- d1 d2) 365.)))

(comment
  (abs -10)
  (abs (flatten [-10]))
  :end)

