(ns functions
  (:require
   [clojure.string :as str]
   [excel :as excel]
   [expressions :as expressions]
   [clojure.math.numeric-tower :as math])
  (:import 
   [java.time LocalDateTime]
   [java.util Calendar Calendar$Builder]))

(defn fn-equal 
  "Replacement for `=` to handle cases where
   the value against which to compare `v1` (`v2`) is
   a regular expression. Used by certain functions
   that allow a comparison to be specified in the
   Excel formula (e.g. SUMIF)
   TODO: check if an RE can be used with a number.
   Currently, this is allowed, but it may not be 
   consistent with how Excel works."
  [v1 v2]
  (if (instance? java.util.regex.Pattern v2)
    (re-matches v2 (str v1))
    (= v1 v2)))

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

(defn fn-search [look-for-str in-str & [starting-at]]
  (cond
    (and starting-at ((complement number?) starting-at))
    excel/VALUE-ERROR
    (and starting-at ((complement pos?) starting-at))
    excel/VALUE-ERROR
    :else
    (if-let [i (str/index-of in-str look-for-str (or (some-> starting-at dec) 0))]
      (-> i (inc) (int))
      excel/VALUE-ERROR)))

(defn fn-sum [& vs]
  (apply + (flatten vs)))

(defn- wrap-if [base-fn]
  (fn [search-range criteria sum-range]
    (let [expr-code (-> criteria
                      (expressions/recast-comparative-expression)
                      (expressions/->code)
                      (expressions/code->with-regex))
        filtered-range (expressions/reduce-by-comp-expression expr-code search-range sum-range)]
    #_(tap> {:loc base-fn
           :search-range search-range
           :sum-range sum-range
           :criteria criteria
           :recast (expressions/recast-comparative-expression criteria)
           :code expr-code
           :f-range filtered-range
           :result (apply base-fn filtered-range)})
    (apply base-fn filtered-range))))

(defn fn-sumif [& [search-range criteria sum-range]]
  ((wrap-if fn-sum) search-range criteria sum-range))

(defn fn-max [& vs]
  (apply max (flatten vs)))

(defn fn-min [& vs]
  (apply min (flatten vs)))

(defn fn-count [& vs]
  (count (keep #(when (number? %) %) vs)))

(defn fn-count-if [& [search-range criteria sum-range]]
  ((wrap-if fn-count) search-range criteria sum-range))

(defn fn-counta [& vs]
  (-> (keep #(when (not (str/blank? (str %))) %) (flatten vs))
      (count)
      (float)))

(defn fn-average [& vs]
  (let [c-vs (flatten vs)]
    (/ (apply fn-sum c-vs)
       (apply fn-count c-vs))))

(defn fn-average-if [& [search-range criteria sum-range]]
  ((wrap-if fn-average) search-range criteria sum-range))

(defn fn-concatenate [& vs]
  (apply str (flatten vs)))

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

(defn- convert-vector-to-table [array-as-vector]
  (if-let [meta-data (meta array-as-vector)]
    (let [{:keys [single? column? cols rows]} meta-data]
      (partition cols array-as-vector))
    array-as-vector))

(defn fn-index [& [lookup-range row-num col-num :as vs]]
  (let [{:keys [rows cols]} (meta lookup-range)
        r-offset (if (= 1 rows) 1 row-num)
        c-offset (if (= 1 rows) row-num (or col-num 1))] 
    (some-> lookup-range
            (convert-vector-to-table)
            (nth (dec r-offset) nil)
            (nth  (dec c-offset) nil))))

(defn fn-vlookup [& [lookup-value table-array-as-vector col-index range-lookup]]
  (let [table-array (convert-vector-to-table table-array-as-vector)
        r-val (some
               (fn [[s-val :as table-row]]
                 (when (= lookup-value s-val)
                   (nth table-row (dec col-index))))
               table-array)]
    (tap> {:loc fn-vlookup
           :lookup lookup-value
           :table-vector table-array-as-vector
           :col-index col-index
           :range-lookup range-lookup
           :table-array (convert-vector-to-table table-array-as-vector)
           :return r-val})
    r-val))

(comment
  (abs -10)
  (abs (flatten [-10]))
  :end)

