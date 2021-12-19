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

(defn- is-multi-range? [lookup-range-or-reference]
  (cond (meta lookup-range-or-reference)
        false
        (every? some? (map meta lookup-range-or-reference))
        true
        :else
        (throw (Exception. IllegalArgumentException "Bad Range"))))

(defn- convert-single-vector-to-table [array-as-vector]
  (if-not (meta array-as-vector)
    (throw (Exception.
            (str "Invalid vector supplied to convert-single-vector-to-table"
                 (pr-str array-as-vector))))
    (->> (meta array-as-vector)
         (:areas)
         (mapv
          (fn [meta-data]
            (let [{:keys [single? column? cols rows]} meta-data]
              (partition cols array-as-vector)))))))

(defn- convert-vector-to-table 
  "Given a vector (or vectors of vectors) of values convert it 
   (or them) to tabular data structure(s) if possible, otherwise
   return the supplied vector unchanged.
   A vector can be converted to a tabular structure if it contains
   meta data describing the shape of the table (i.e. number of rows,
   cols etc.)

   Example meta data might look like:

     {:areas [{:single? false
               :column? false
               :sheet-name \"Sheet5\"
               :tl-name \"A1\"
               :tl-coord [0 0]
               :cols 3
               :rows 4}]}
   
   and an example return value might be:

     [((\"Cashews\" 3.55 16.0)
       (\"Peanuts\" 1.25 20.0)
       (\"Walnuts\" 1.75 12.0)
       (nil nil nil))]"
  [array-as-vector]
  (cond (meta array-as-vector) ; a single vector with metadata
        (convert-single-vector-to-table array-as-vector)
        (every? some? (map meta array-as-vector)) ; a vector of vectors with metadata
        (map 
         (fn [array-as-vector-component]
           (convert-single-vector-to-table array-as-vector-component))
         array-as-vector)
        :else  ; an unadorned vector
        array-as-vector))

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
  ;;   because we only have one area to concern ourselves with and
  ;;   even though the :areas content of the metadata is a vector
  ;;   we only consider the first area, which should be the only one.
  ;; If this is a reference call
  ;;   we need to pay attention to the area
  (let [is-multi? (is-multi-range? lookup-range-or-reference)
        {:keys [rows cols]}
        (if is-multi?
          (-> (map meta lookup-range-or-reference) 
              (nth (or (some-> area-num dec int) 0))
              (:areas)
              (first))
          (-> lookup-range-or-reference 
              (meta) 
              (:areas) 
              (first)))
        r-offset (if (= 1 rows) 1 row-num)
        c-offset (if (= 1 rows) row-num (or col-num 1))
        a-offset (if area-num (dec area-num) 0)
        table (cond->
               (convert-vector-to-table lookup-range-or-reference)
                is-multi?
                (nth a-offset))]
    (tap> {:lookup-range lookup-range-or-reference
           :multi? is-multi?
           :meta (if is-multi?
                   (map meta lookup-range-or-reference)
                   (meta lookup-range-or-reference))
           :convert table})
    (some-> table
            (first)
            (extract-from-table-view r-offset c-offset))))

(comment
  (fn-index ^{:areas [{:cols 3}]} ["11" "12" "13" "21" "22" "23"] 2 3)
  (some-> '(["Fruit" "Price" "Count" "Apples" 0.69 40.0 "Bananas" 0.34 38.0 "Lemons" 0.55 15.0 "Oranges" 0.25 25.0 "Pears" 0.59 40.0]
            ["Cashews" 3.55 16.0 "Peanuts" 1.25 20.0 "Walnuts" 1.75 12.0 nil nil nil])
          (nth 1.0 nil)
          (convert-vector-to-table)
          #_(nth 1.0 nil)
          #_(nth 1.0 nil))
  
  (convert-vector-to-table
   [1 2 3 4])
  
  (convert-vector-to-table
   ^{:areas [{:cols 2}]}
   [1 2 3 4])

  (convert-vector-to-table
   [^{:areas [{:cols 2}]} [1 2 3 4]
    ^{:areas [{:cols 1}]} [1 2 3 4]])
  
  :end)

(defn fn-index-reference [& [lookup-range row-num col-num :as vs]]
  (let [{:keys [rows cols tl-coord]} (-> lookup-range (meta) :areas (first))
        [tl-row tl-col] tl-coord 
        r-offset (if (= 1 rows) 1 row-num)
        c-offset (if (= 1 rows) row-num (or col-num 1))]
    (str
     (CellReference/convertNumToColString (int (+ tl-col (dec c-offset))))
     (int (+ (inc tl-row) (dec r-offset))))))

(defn- is-properly-sorted? [range-vec & [ordering]]
  (let [compare-fail-fn
        (let [ordering-int (or (some-> ordering (int)) 1)]
          (if (zero? ordering-int)
            (constantly false) ; never fail when we say we have an unsorted range
            (fn [v1 v2] ; function to decide if we fail sorted range test
              (= (case ordering-int
                   1 1
                   -1 -1)
                 (compare v1 v2)))))]
    (loop [v range-vec last-val nil shortcut-exit false]
      (cond shortcut-exit ; found a value that wasn't sorted as expected
            false
            (not (seq v)) ; done with the seq and no ill-sorted items found
            true
            :else
            (let [val (first v)]
              (recur
               (rest v)
               val
               ;; fail if we find an ill-sorted item
               (and last-val (compare-fail-fn last-val val))))))))

(comment
  (is-properly-sorted? [1 1 2 3 3 3 3 4 4] 1)
  (is-properly-sorted? ["A" "B" "C"] 1)
  (is-properly-sorted? ["C" "B" "B" "A"] 0)
  (is-properly-sorted? [4 4 3 2 1] -1))

(defn fn-match [& [lookup-val lookup-vec match-type :as vs]]
  (let [properly-sorted? (is-properly-sorted? lookup-vec match-type)] 
    (if properly-sorted?
      (loop [v lookup-vec pos 0 prev-canditate? nil r nil]
        (cond
          (or (and (zero? match-type) (some? prev-canditate?) (true? prev-canditate?))
              (and (not (zero? match-type)) (some? prev-canditate?) (false? prev-canditate?)))
          r
          (not (seq v))
          (if prev-canditate?
            r
            "#N/A")
          :else
          (let [c-val (first v)
                is-canditate? 
                (case match-type
                  0 (if (or (string? c-val)
                            (string? lookup-val))
                      (= 0 (compare (-> c-val (str)) (-> lookup-val (str))))
                      (= 0 (compare c-val lookup-val)))
                  (not= match-type (compare c-val lookup-val)))]
            (println c-val is-canditate? pos)
            (recur
             (rest v)
             (inc pos)
             is-canditate?
             (if is-canditate? pos r)))))
      "#ERROR")))

(comment 
  (fn-match 6 [1 3 5 7 9 11 13 15] 1)
  (fn-match 15 [1 3 5 7 9 11 13 15] 1)
  (fn-match 4 [15 13 11 9 7 5 3 1] -1)
  (fn-match 4 [9 7 5 4 3 1] 0)
  (fn-match 90 [9 7 5 4 3 1] 0)
  (fn-match "B*" ["A" "B" "C"] 0)

  (let [test-expr "A~**~?"
        comp-string "A*B?"] 
    (let
     [looks-like-wc? (re-matches #".*((?<!~)\*|(?<!~)\?).*" test-expr)
      re (if looks-like-wc?
           (-> test-expr
               (str/replace
                #"([^~])(\*)"
                "$1.*")
               (str/replace
                #"([^~])(\?)"
                "$1.")
               (str/replace
                #"(~\*)"
                "\\\\*")
               (str/replace
                #"(~\?)"
                "\\\\?")
               (re-pattern))
           (-> test-expr
               (str/replace
                #"(~\*)"
                "\\\\*")
               (str/replace
                #"(~\?)"
                "\\\\?")
               (re-pattern)))]
      [looks-like-wc? re (re-matches re comp-string)]))

  (-> "A*"
      (str/replace
       #"(~\*)"
       "*")
      (str/replace
       #"(~\?)"
       "?")
      (pr-str)
      (re-pattern))

  (re-matches (re-pattern "A.*") "ABC"))

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

