(ns functions
  (:require
   [clojure.string :as str]
   [excel :as excel]
   [clojure.math.numeric-tower :as math])
  (:import 
   [java.time LocalDateTime]
   [java.util Calendar Calendar$Builder]))

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

(defn fn-date [& [year month day]]
  (let [cal (excel/build-calendar-for-year-and-advance 
              (if (<= 0 year 1899) (+ 1900 year) year) month day)
        tz-id (-> cal
                  (.getTimeZone)
                  (.toZoneId))]
    (->
     (LocalDateTime/ofInstant
      (.toInstant cal)
      tz-id)
     (excel/local-date-time->excel-serial-date))))

(comment
  (excel/build-calendar-for-year-and-advance 2020 1 15)
  (excel/build-calendar-for-year-and-advance 2019 14 29)
  (excel/build-calendar-for-year-and-advance 2020 14 29)
  (excel/build-calendar-for-year-and-advance 2021 14 29)
  (excel/build-calendar-for-year-and-advance 2021 14 -1)
  (excel/build-calendar-for-year-and-advance 2021 -3 -1)
  (fn-date 2020 1 15)
  :end)

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

