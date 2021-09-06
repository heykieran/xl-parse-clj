(ns functions
  (:require
   [clojure.string :as str]
   [excel :as excel]
   [clojure.math.numeric-tower :as math]))

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

(defn fn-datevalue [v]
  (excel/parse-excel-string-to-serial-date v))

(defn fn-yearfrac [& [date-1 date-2 b]]
  (let [d1 (if (number? date-1)
             date-1
             (excel/parse-excel-string-to-serial-date date-1))
        d2 (if (number? date-2)
             date-2
             (excel/parse-excel-string-to-serial-date date-2))]
    (case b
      (nil 0.) (excel/nasd-360-diff d1 d2)
      1. (excel/act-act-diff d1 d2)
      2. (math/abs (/ (- d1 d2) 360.))
      3. (math/abs (/ (- d1 d2) 365.))
      4. (excel/euro-360-diff d1 d2))))

(comment
  (abs -10)
  (abs (flatten [-10]))
  :end)

