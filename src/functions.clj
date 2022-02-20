(ns functions
  (:require
   [clojure.string :as str]
   [excel :as excel]
   [graph :as graph]
   [expressions :as expressions]
   [clojure.math.numeric-tower :as math])
  (:import
   [java.time LocalDateTime ZoneId]
   [java.util Calendar Calendar$Builder]
   [org.apache.poi.ss.util CellReference]
   [org.apache.poi.ss.formula.functions Finance]
   [box Box]))

(defn- unbox-value 
  "If v is a boxed value unbox it else return v"
  [v]
  (if (instance? Box v) @v v))

(defn- unbox-seq
  "If v is a boxed seq of values unbox then else return v"
  [v]
  (let [v-value (unbox-value v)]
    (map unbox-value v-value)))

(defn fn-equal?
  "Replacement for `=` to handle cases where
   the value against which to compare `v1` (`v2`) is
   a regular expression. Used by certain functions
   that allow a comparison to be specified in the
   Excel formula (e.g. SUMIF)
   TODO: check if an RE can be used with a number.
   Currently, this is allowed, but it may not be 
   consistent with how Excel works."
  [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (if (instance? java.util.regex.Pattern v2-value)
      (re-matches v2-value (str v1-value))
      (= v1-value v2-value))))

(defn fn-unary-plus [v1]
  (let [v1-value (unbox-value v1)]
    (+ v1-value)))

(defn fn-unary-minus [v1]
  (let [v1-value (unbox-value v1)]
    (- v1-value)))

(defn fn-add [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (+ v1-value v2-value)))

(defn fn-subtract [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (- v1-value v2-value)))

(defn fn-multiply [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (* v1-value v2-value)))

(defn fn-divide [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (/ v1-value v2-value)))

(defn fn-exponent [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (math/expt v1-value v2-value)))

(defn fn-gt? [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (> v1-value v2-value)))

(defn fn-lt? [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (< v1-value v2-value)))

(defn fn-gt-equal? [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (>= v1-value v2-value)))

(defn fn-lt-equal? [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (<= v1-value v2-value)))

(defn fn-not-equal? [v1 v2]
  (let [v1-value (unbox-value v1)
        v2-value (unbox-value v2)]
    (not= v1-value v2-value)))

(defn fn-abs [v]
  (let [v1-value (unbox-value v)]
    (if (neg? v1-value)
      (- v1-value)
      v1-value)))

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

(defn fn-prcnt [v]
  (/ (bigdec v) 100.0M))

(defn pi []
  (Math/PI))

(defn fn-concat [& vs]
  (apply str (map #(if (instance? Box %) @% %) vs)))

(defn fn-mod [n d]
  (mod n d))

(defn fn-sign [^double n]
  (cond (zero? n)
        0
        (neg? n)
        -1
        :else
        1))

(defn fn-round 
  ([^double n ^long p]
   (fn-round n p java.math.RoundingMode/HALF_UP))
  ([^double n ^long p ^long rounding-mode]
  (if (or (Double/isNaN n)
          (Double/isInfinite n))
    Double/NaN
    (-> n
        (java.math.BigDecimal/valueOf)
        (.setScale p rounding-mode)
        (.doubleValue)))))

(defn fn-roundup [^double n ^long p]
  (fn-round n p java.math.RoundingMode/UP))

(defn fn-rounddown [^double n ^long p]
  (fn-round n p java.math.RoundingMode/DOWN))

(defn fn-floor [^double n ^double s]
  (if (and (zero? s) (not (zero? n)))
    Double/NaN
    (if (or (zero? n)
            (zero? s))
      0
      (* s (Math/floor (/ n s))))))

(defn fn-ceiling [^double n ^double s]
  (if (and (pos? n) (neg? s))
    Double/NaN
    (if (or (zero? n)
            (zero? s))
      0
      (* s (Math/ceil (/ n s))))))

(defn fn-pmt 
  ([rate nper pv]
   (fn-pmt rate nper pv 0 0))
  ([rate nper pv fv]
   (fn-pmt rate nper pv fv 0))
  ([rate nper pv fv c-type]
   (Finance/pmt rate nper pv (or fv 0) (or c-type 0))))

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
  (let [values (flatten (unbox-seq vs))]
    (apply + (filter number? values))))

(defn- wrap-if [base-fn]
  (fn [search-range criteria sum-range]
    (let [expr-code (-> criteria
                        ((fn [c] (if (instance? Box c) @c c)))
                        (expressions/recast-comparative-expression)
                        (expressions/->code)
                        (expressions/code->with-regex))
          filtered-range (expressions/reduce-by-comp-expression
                          expr-code
                          search-range
                          sum-range)]
      (apply base-fn filtered-range))))

(defn fn-sumif [& [search-range criteria sum-range]]
  ((wrap-if fn-sum) search-range criteria sum-range))

(defn fn-max [& vs]
  (apply max (->> vs unbox-seq flatten (remove nil?))))

(defn fn-min [& vs]
  (apply min (->> vs unbox-seq flatten (remove nil?))))

(defn fn-count [& vs]
  (let [values (flatten (unbox-seq vs))]
    (->> values
         (keep #(when (number? %) %))
         (count))))

(defn fn-count-if [& [search-range criteria sum-range]]
  ((wrap-if fn-count) search-range criteria sum-range))

(defn fn-counta [& vs]
  (let [values (flatten (unbox-seq vs))]
    (-> (keep #(when (not (str/blank? (str %))) %) values)
        (count)
        (float))))

(defn fn-average [& vs]
  (/ (apply fn-sum vs)
     (apply fn-count vs)))

(defn fn-average-if [& [search-range criteria sum-range]]
  ((wrap-if fn-average) search-range criteria sum-range))

(defn fn-concatenate [& vs]
  (let [values (flatten (unbox-seq vs))]
    (apply str values)))

(defn fn-now []
  (excel/excel-now))

(defn fn-date [& [year month day]]
  (let [cal (excel/build-calendar-for-year-and-advance
             (if (<= 0 year 1899)
               (+ 1900 year)
               year)
             month day)
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
  (fn-subtract d1 d2))

(defn fn-datevalue [v]
  (excel/parse-excel-string-to-serial-date (unbox-value v)))

(defn fn-yearfrac [& [date-1 date-2 b]]
  (let [date-1-value (unbox-value date-1)
        date-2-value (unbox-value date-2)
        d1 (if (number? date-1-value)
             date-1-value
             (excel/parse-excel-string-to-serial-date date-1-value))
        d2 (if (number? date-2-value)
             date-2-value
             (excel/parse-excel-string-to-serial-date date-2-value))]
    (case b
      (nil 0.) (excel/nasd-360-diff d1 d2)
      1. (excel/act-act-diff d1 d2)
      2. (math/abs (/ (- d1 d2) 360.))
      3. (math/abs (/ (- d1 d2) 365.))
      4. (excel/euro-360-diff d1 d2))))

(defn fn-year [date-serial]
  (-> date-serial
      (excel/build-calendar-for-serial-date)
      (excel/extract-date-fields)
      (nth 0)))

(defn fn-month [date-serial]
  (-> date-serial
      (excel/build-calendar-for-serial-date)
      (excel/extract-date-fields)
      (nth 1)))

(defn fn-day [date-serial]
  (-> date-serial
      (excel/build-calendar-for-serial-date)
      (excel/extract-date-fields)
      (nth 2)))

(defn fn-eomonth [date-serial months]
  (excel/advance-and-get-end-of-month date-serial months))

(defn fn-edate [date-serial months]
  (excel/advance-and-get-date date-serial months))

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
              (partition cols @array-as-vector)))))))

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
  (tap> {:loc fn-index
         :lookup-range lookup-range-or-reference
         :is-multi? (is-multi-range? lookup-range-or-reference)})
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
    #_(tap> {:lookup-range lookup-range-or-reference
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

(defn text-equal-with-possible-wildcard?
  "Compare two strings to see if they're equal. The test-expr
   may contain wildchars as understood by Excel i.e. * and ?, 
   where either can be escaped by preceding them with a ~. comp-str
   is a string. The comparison is done in a case-insensitive 
   manner."
  [test-expr comp-string]
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
             ((fn [s] (str "(?i)" s)))
             (re-pattern))
         (-> (str "^" test-expr "$")
             (str/replace
              #"(~\*)"
              "\\\\*")
             (str/replace
              #"(~\?)"
              "\\\\?")
             ((fn [s] (str "(?i)" s)))
             (re-pattern)))]
    (some-> comp-string
            (fn-equal? re)
            (some?)
            (not= false))))

(comment
  (text-equal-with-possible-wildcard? "B~*A" "B*A")
  (text-equal-with-possible-wildcard? "B~*A" "B*C")
  (text-equal-with-possible-wildcard? "B*A" "BBBBBA")
  (text-equal-with-possible-wildcard? "BA*" "BA")
  (text-equal-with-possible-wildcard? "BA?" "BAC")
  (text-equal-with-possible-wildcard? "BA?" "BACD")
  (text-equal-with-possible-wildcard? "BA*" "BACD")
  :end)

(defn fn-match [& [lookup-val lookup-vec match-type-any :as vs]]
  (let [match-type (or (some-> match-type-any (int)) 1)
        lookup-vec-values (if (instance? Box lookup-vec) @lookup-vec lookup-vec)
        properly-sorted? (is-properly-sorted? lookup-vec-values match-type)]
    (if properly-sorted?
      (loop [v lookup-vec-values pos 0 prev-canditate? nil r nil]
        (cond
          (or (and (zero? match-type) (some? prev-canditate?) (true? prev-canditate?))
              (and (not (zero? match-type)) (some? prev-canditate?) (false? prev-canditate?)))
          (some-> r (inc))
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
                      (text-equal-with-possible-wildcard?
                       (-> lookup-val (str))
                       (-> c-val (str)))
                      (= 0 (compare c-val lookup-val)))
                  (not= match-type (compare c-val lookup-val)))]
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
  (fn-match "B*" ["ABC" "BC" "C"] 0)
  (fn-match "B~*A" ["ABC" "BC" "B*A" "C"] 0)
  (fn-match "B*" ["ABC" "BC" "C"] 0)
  (fn-match 41.0 [25.0 38.0 40.0 41.0] 0.0)
  :end)

(defn fn-indirect [& [sheet-name context ref-text a1]]
  (let [value (if (instance? Box ref-text) @ref-text ref-text)
        [_ sheet-with-exclam cell] (re-matches #"(.*!)?(.*)" value)]
    (if sheet-with-exclam
      value
      (str sheet-name "!" value))))

(defn fn-offset [& [sheet-name context reference rows cols height width]]
  (let [base-cell (-> reference meta :areas (first) :tl-name)
        target-cell (when base-cell
                      (excel/ref-str->ref-str-using-offset
                       base-cell
                       rows cols))
        target-cell-2 (when base-cell
                        (excel/ref-str->ref-str-using-offset
                         base-cell
                         (+ rows (dec (or height 1))) (+ cols (dec (or width 1)))))
        target-value (when (and (not= excel/REF-ERROR target-cell)
                                (not= excel/REF-ERROR target-cell-2))
                       (if (= target-cell target-cell-2)
                         (graph/eval-range (str sheet-name "!" target-cell) context)
                         (graph/eval-range (str sheet-name "!" target-cell ":" target-cell-2) context)))]
    (tap> {:loc fn-offset
           :sheet-name sheet-name
           :meta (meta reference)
           :base-cell base-cell
           :target-cell target-cell
           :target-cell-2 target-cell-2
           :target-value target-value
           :rows rows
           :cols cols
           :height height
           :width width})
    target-value))

(comment (excel/ref-str->ref-str-using-offset "D3" -3 -3))

(defn fn-vlookup [& [lookup-value table-array-as-vector col-index range-lookup]]
  (assert (= 1 (-> table-array-as-vector (meta) :areas (count)))
          "Only expecting one area in meta data for fn-vlookup")
  (let [table-array (-> table-array-as-vector convert-vector-to-table (first))
        r-val (some
               (fn [[s-val :as table-row]]
                 (when (fn-equal? lookup-value s-val)
                   (nth table-row (dec col-index))))
               table-array)]
    #_(tap> {:loc fn-vlookup
             :lookup lookup-value
             :table-vector table-array-as-vector
             :col-index col-index
             :range-lookup range-lookup
             :table-array table-array
             :meta-table-array (meta table-array-as-vector)
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
    #_(tap> {:loc fn-union
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

