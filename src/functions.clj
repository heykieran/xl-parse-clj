(ns functions
  (:require
   [clojure.string :as str]
   [excel :as excel]
   [expressions :as expressions]
   [clojure.math.numeric-tower :as math])
  (:import 
   [java.time LocalDateTime]
   [java.util Calendar Calendar$Builder]
   [org.apache.poi.ss.util CellReference]))

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
  (apply + (filter number? (flatten vs))))

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
  (if-let [meta-data-map (meta array-as-vector)]
    (->> (:areas meta-data-map)
         (mapv (fn [meta-data]
                 (let [{:keys [single? column? cols rows]} meta-data]
                   (partition cols array-as-vector)))))
    array-as-vector))

(defn- is-multi-range? [lookup-range-or-reference]
  (cond (meta lookup-range-or-reference)
        false
        (every? some? (map meta lookup-range-or-reference))
        true
        :else
        (throw (Exception. IllegalArgumentException "Bad Range"))))

(defn- extract-from-table-view [as-table r-offset c-offset]
  (tap> {:as-table as-table
         :r-offset r-offset
         :c-offset c-offset})
  (letfn [(not-zero? [n] ((complement zero?) n))]
    (cond (and (not-zero? r-offset) (not-zero? c-offset))
        (-> as-table
            (nth (dec r-offset) nil)
            (nth (dec c-offset) nil))
        (and (not-zero? r-offset) (zero? c-offset))
        (-> as-table
            (nth (dec r-offset) nil))
        (and (zero? r-offset) (not-zero? c-offset))
        (->> as-table
             (map #(nth % (dec c-offset) nil)))
        :else
        as-table)))

(defn fn-index [& [lookup-range-or-reference row-num col-num area-num :as vs]]
  ;; If this is an array call
  ;;   because we only have one area to convern ourselves with and
  ;;   even though the :areas content of the metadata is a vector
  ;;   we only consider the first area, which should be the only one.
  ;; If this is a reference call
  ;;   we need to pay attention to the area
  (let [{:keys [rows cols]}
        (if (is-multi-range? lookup-range-or-reference)
          (-> lookup-range-or-reference (meta) :areas (nth (or (some-> area-num dec int) 0)))
          (-> lookup-range-or-reference (meta) :areas (first)))
        r-offset (if (= 1 rows) 1 row-num)
        c-offset (if (= 1 rows) row-num (or col-num 1))]
    (tap> {:lookup-range lookup-range-or-reference
           :multi? (is-multi-range? lookup-range-or-reference)
           :meta (meta lookup-range-or-reference)
           :r-offset r-offset
           :c-offser c-offset
           :table (-> lookup-range-or-reference
                      (convert-vector-to-table))})
    (some-> lookup-range-or-reference
            (convert-vector-to-table)
            ((fn [t]
               (if (nil? area-num)
                 (assert (= 1 (count t)) 
                         (str "Only expecting one area. "
                                                  "Count = " (count t) ", "
                                                  "t = " (pr-str t)))
                 (assert (>= (count t) area-num)
                         (str "Area " area-num " requested, but not available "
                              "Count = " (count t) ", "
                              "t = " (pr-str t))))
               #_(tap> {:t t
                      :meta (mapv meta t)
                      :m (mapv convert-vector-to-table t)
                      :target (nth (map convert-vector-to-table t) (or (some-> area-num dec int) 0))
                      :v (some-> 
                          (mapcat convert-vector-to-table t)
                          (nth (or (some-> area-num dec int) 0))
                          (nth (dec r-offset) nil)
                          (nth (dec c-offset) nil)
                          )
                      :a (or (some-> area-num dec) 0)
                      :r r-offset
                      :c c-offset})
               t))
            (first)
            #_(nth (or (some-> area-num dec int) 0))
            (extract-from-table-view r-offset c-offset))))

(comment
  (fn-index ^{:areas [{:cols 3}]} ["11" "12" "13" "21" "22" "23"] 2 3)
  (some-> '(["Fruit" "Price" "Count" "Apples" 0.69 40.0 "Bananas" 0.34 38.0 "Lemons" 0.55 15.0 "Oranges" 0.25 25.0 "Pears" 0.59 40.0]
    ["Cashews" 3.55 16.0 "Peanuts" 1.25 20.0 "Walnuts" 1.75 12.0 nil nil nil])
      (nth 1.0 nil)
          (convert-vector-to-table)
      #_(nth 1.0 nil)
      #_(nth 1.0 nil))
  )

(defn fn-index-reference [& [lookup-range row-num col-num :as vs]]
  (let [{:keys [rows cols tl-coord]} (-> lookup-range (meta) :areas (first))
        [tl-row tl-col] tl-coord 
        r-offset (if (= 1 rows) 1 row-num)
        c-offset (if (= 1 rows) row-num (or col-num 1))]
    (str
     (CellReference/convertNumToColString (int (+ tl-col (dec c-offset))))
     (int (+ (inc tl-row) (dec r-offset))))))

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

(defn fn-union [& vs]
  (let [result (loop [vs-to-combine vs combined []]
                 (if-not (seq vs-to-combine)
                   combined
                   (let [curr-vs (first vs-to-combine)
                         {:keys [tl-name tl-coord rows cols] :as curr-meta} (meta curr-vs)
                         curr-table (convert-vector-to-table curr-vs)]
                     
                     (recur (rest vs-to-combine)
                            (concat combined curr-table)))))]
    (tap> {:loc fn-union
           :result result
           :vs (mapv (fn [v]
                       {:v v
                        :meta (meta v)
                        :t (convert-vector-to-table v)})
                     vs)}))
  vs)



(defn fn-range [& [:as vs]]
  vs)

(comment
  (abs -10)
  (abs (flatten [-10]))
  :end)

