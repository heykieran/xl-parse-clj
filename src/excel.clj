(ns excel
  (:require
   [clojure.string :as str]
   [dk.ative.docjure.spreadsheet :as dk]
   [clojure.math.numeric-tower :as math])
  (:import
   [java.util Locale TimeZone Calendar Calendar$Builder]
   [java.time LocalDate LocalTime LocalDateTime]
   [java.time.format DateTimeFormatter]
   [java.text DateFormatSymbols]
   [org.apache.poi.util LocaleUtil]
   [org.apache.poi.ss.usermodel CellType DateUtil]))

(def PATTERNS
  [[:YMD-DASHES {:r #"^(\d{4})-(\w+)-(\d{1,2})( .*)?$", :s "ymd" :m-type :num}]
   [:DMY-DASHES {:r #"^(\d{1,2})-(\w+)-(\d{4})( .*)?$", :s "dmy" :m-type :num}]
   [:MD-DASHES {:r #"^(\w+)-(\d{1,2})( .*)?$", :s "md" :m-type :num}]
   [:MDY-SLASHES {:r #"^(\w+)/(\d{1,2})/(\d{4})( .*)?$", :s "mdy" :m-type :num}]
   [:YMD-SLASHES {:r #"^(\d{4})/(\w+)/(\d{1,2})( .*)?$", :s "ymd" :m-type :num}]
   [:MD-SLASHES {:r #"^(\w+)/(\d{1,2})( .*)?$", :s "md" :m-type :num}]])

(defn generate-month-symbol-lookup []
  (reduce (fn [accum inst-fn]
            (as-> (DateFormatSymbols/getInstance) $
              (inst-fn $)
              (remove #(str/blank? %) $)
              (map-indexed (fn [idx m-str] [(str/lower-case m-str) (inc idx)]) $)
              (apply conj accum $)))
          {}
          [(memfn getShortMonths) (memfn getMonths)]))

(def month-symbol-lookup
  (delay (generate-month-symbol-lookup)))

(defn get-symbol-match-string [m-type]
  (let [f (case m-type
            :month (memfn getMonths)
            :short-month (memfn getShortMonths))]
    (as->
     (DateFormatSymbols/getInstance) $
      (f $)
      (remove #(str/blank? %) $)
      (str/join "|" $))))

(def LONG-PATTERNS
  (let [syms-short (get-symbol-match-string :short-month)
        syms-long (get-symbol-match-string :month)]
    [[:MDY-LONG {:r (re-pattern (str "(?i)^(" syms-long ")\\s*(\\d{1,2})\\s*,\\s{1,}(\\d{4})(. *)?$"))
                 :s "mdy" :m-type :string}]
     [:DMY-LONG {:r (re-pattern (str "(?i)^(\\d{1,2})\\s*(" syms-long ")\\s*(\\d{2,})(. *)?$"))
                 :s "dmy" :m-type :string}]
     [:MD-LONG {:r (re-pattern (str "(?i)^(" syms-long ")(?:\\s*|-*)(\\d{1,2})(. *)?$"))
                :s "md" :m-type :string}]
     [:MY-LONG {:r (re-pattern (str "(?i)^(" syms-long ")\\s*(\\d{4})(. *)?$"))
                :s "my" :m-type :string}]
     [:MDY-SHORT {:r (re-pattern (str "(?i)^(" syms-short ")\\s*(\\d{1,2})\\s*,\\s{1,}(\\d{4})(. *)?$"))
                  :s "mdy" :m-type :string}]
     [:DMY-SHORT {:r (re-pattern (str "(?i)^(\\d{1,2})\\s*(" syms-short ")\\s*(\\d{2,})(. *)?$"))
                  :s "dmy" :m-type :string}]
     [:MD-SHORT {:r (re-pattern (str "(?i)^(" syms-short ")(?:\\s*|-*)(\\d{1,2})(. *)?$"))
                 :s "md" :m-type :string}]
     [:MY-SHORT {:r (re-pattern (str "(?i)^(" syms-short ")\\s*(\\d{4})(. *)?$"))
                 :s "my" :m-type :string}]]))

(defn parse-excel-string-to-date-info [date-string]
  (letfn [(cvt-int [v] (if (and v (string? v))
                         (Integer/parseInt v)
                         v))]
    (some
     (fn [[k {:keys [r s m-type]}]]
       (when-let [[_ v1 v2 v3] (re-matches r date-string)]
         (let [vs [v1 v2 v3]
               o-y (str/index-of s "y")
               o-m (str/index-of s "m")
               o-d (str/index-of s "d")
               y-n (let [current-y (-> (LocaleUtil/getUserTimeZone)
                                       (.toZoneId)
                                       (LocalDate/now)
                                       (.getYear))]
                     (if (nil? o-y)
                       current-y
                       (let [y-part (nth vs o-y)]
                         (cond
                           (= 4 (count y-part))
                           (cvt-int y-part)
                           (= 2 (count y-part))
                           (+ (* 100 (quot current-y 100))
                              (cvt-int y-part))
                           :else
                           (throw (IllegalArgumentException.
                                   (str "Can't convert year string " o-y)))))))
               m-n (if (= :num m-type)
                     (cvt-int (nth vs o-m))
                     (get (deref month-symbol-lookup)
                          (str/lower-case (nth vs o-m))))
               d-n (if o-d
                     (cvt-int (nth vs o-d))
                     1)]
           {:pattern-name k
            :pattern-style s
            :date (LocalDate/of y-n m-n d-n)})))
     (concat LONG-PATTERNS PATTERNS))))

(defn parse-excel-string-to-serial-date [date-str]
  (if-let [{:keys [pattern-name pattern-style date]}
           (excel/parse-excel-string-to-date-info date-str)]
    (excel/local-date-time->excel-serial-date date)
    "#VALUE!"))

(comment
  (ns-unmap *ns* 'month-symbol-lookup)
  @month-symbol-lookup
  (parse-excel-string-to-date-info "1/15/2021")
  (parse-excel-string-to-date-info "2021/1/15")
  (parse-excel-string-to-date-info "January 15, 2021")
  (parse-excel-string-to-date-info "Jan 2021")
  (parse-excel-string-to-date-info "1/15")
  (parse-excel-string-to-date-info "Jan 15")
  (parse-excel-string-to-date-info "Jan-15")
  (parse-excel-string-to-date-info "Jan 15, 2021")
  (parse-excel-string-to-date-info "january 15, 2021")
  (parse-excel-string-to-date-info "Jan2021")
  (parse-excel-string-to-date-info "15 Jan 21")
  :end)

(defn excel-serial-date->local-date-time-manual
  "Convert an Excel serial date to a local date time instance"
  [excel-serial-date]
  (letfn
   [(s->EpochDays [s]
      (+ -25568 (if (> s 59) (dec s) s)))]
    (let [s-d (long excel-serial-date)
          s-f (Math/round (* (- excel-serial-date s-d) 24 60 60))]
      (LocalDateTime/of
       (LocalDate/ofEpochDay (s->EpochDays s-d))
       (LocalTime/ofSecondOfDay s-f)))))

(defn excel-serial-date->local-date-time
  "Convert an Excel serial date to a local date time instance"
  [excel-serial-date]
  (DateUtil/getJavaDate
   excel-serial-date
   (TimeZone/getTimeZone "UTC")))

(defn string->local-date-time
  "Convert a string to a local datetime object"
  ([ld-str]
   (string->local-date-time ld-str DateTimeFormatter/ISO_LOCAL_DATE_TIME))
  ([ld-str dt-formatter]
   (LocalDateTime/parse
    ld-str dt-formatter)))

(defn local-date-time->excel-serial-date
  "Convert a local date time to an Excel serial date"
  [ldt]
  (DateUtil/getExcelDate
   ldt))

(defn excel-date-fmt->fmt
  "Convert Excel format strings to ones understood by DateFormatter"
  [excel-date-fmt]
  (str/replace
   excel-date-fmt
   #"AM/PM"
   "a"))

(defn local-date-time->string
  "Convert a LocalDateTime instance to a string
   using the format string provided."
  ([ldt]
   (local-date-time->string ldt "yyyy-MM-dd HH:mm"))
  ([ldt format-string]
   (.format ldt
            (-> format-string
                (excel-date-fmt->fmt)
                (DateTimeFormatter/ofPattern)))))

(defn excel-now []
  (-> (LocalDateTime/now)
      (local-date-time->excel-serial-date)))

(defn build-calendar
  "Create a GregorianCalendar object, initialized to the Excel
   serial date supplied, and return it. The serial-date is expected 
   to represent a UTC datetime."
  [serial-date]
  (.. (Calendar$Builder.)
      (setCalendarType "iso8601")
      (setLocale Locale/US)
      (setTimeZone (TimeZone/getTimeZone "UTC"))
      (setInstant (excel/excel-serial-date->local-date-time serial-date))
      (build)))

(defn extract-date-fields
  "Given an initialized calendar instance (containing a UTC instant),
   return a 7-tuple containing the year, month, day, hour, minute, 
   second and millisecond for that instant. Note: for the month field
   January=1 and December=12"
  [calendar-instant]
  (mapv #(cond->
          (.get calendar-instant %)
           (= Calendar/MONTH %)
           (inc))
        [Calendar/YEAR Calendar/MONTH Calendar/DAY_OF_MONTH
         Calendar/HOUR_OF_DAY Calendar/MINUTE Calendar/SECOND
         Calendar/MILLISECOND]))

(defn date-vecs->nasd-date
  "Given two vectors containing year, month and day representing a start date and
   an end date, return a 3-vector containing the NASD modified start and end 
   dates and the number of NASD days, as calculated by NASD 30/360, between
   those two dates"
  [[year-s month-s day-s]
   [year-e month-e day-e]]
  (let [leap-year? (fn [y]
                     (let [d (if (zero? (mod y 100))
                               400
                               4)]
                       (zero? (mod y d))))
        ld-feb-s (if (leap-year? year-s) 29 28)
        ld-feb-e (if (leap-year? year-e) 29 28)
        eff-day-e-i
        (if
         (and (= 2 month-s)
              (= ld-feb-s day-s)
              (= 2 month-e)
              (= ld-feb-e day-e))
          30
          day-e)
        eff-day-s
        (if (or
             (= 31 day-s)
             (and (= 2 month-s) (= ld-feb-e day-s)))
          30
          day-s)
        eff-day-e
        (if (and
             (= 30 eff-day-s)
             (= 31 eff-day-e-i))
          30
          eff-day-e-i)]
    [[year-s month-s eff-day-s]
     [year-e month-e eff-day-e]
     (+ (* (- year-e year-s) 360)
        (* (- month-e month-s) 30)
        (- eff-day-e eff-day-s))]))

(defn date-vecs->euro-date
  "Given two vectors containing year, month and day representing a start date and
   an end date, return a 3-vector containing the Euro360 modified start and end 
   dates and the number of Euro days, as calculated by Euro 30/360, between
   those two dates"
  [[year-s month-s day-s]
   [year-e month-e day-e]]
  (let [eff-day-s
        (if (= 31 day-s) 30 day-s)
        eff-day-e
        (if (= 31 day-e) 30 day-e)]
    [[year-s month-s eff-day-s]
     [year-e month-e eff-day-e]
     (+ (* (- year-e year-s) 360)
        (* (- month-e month-s) 30)
        (- eff-day-e eff-day-s))]))

(defn nasd-360-diff
  "Given two dates, in Excel Serial format, return the number of years as a
   double between those two date,s calculated using the NASD 30/360 methodology."
  [excel-serial-start excel-serial-end]
  (->
   (date-vecs->nasd-date
    (-> (build-calendar excel-serial-start)
        (extract-date-fields))
    (-> (build-calendar excel-serial-end)
        (extract-date-fields)))
   (nth 2)
   (/ 360)
   (double)
   (math/abs)))

(defn euro-360-diff
  "Given two dates, in Excel Serial format, return the number of years as a
   double between those two date,s calculated using the Euro 30/360 methodology."
  [excel-serial-start excel-serial-end]
  (->
   (date-vecs->euro-date
    (-> (build-calendar excel-serial-start)
        (extract-date-fields))
    (-> (build-calendar excel-serial-end)
        (extract-date-fields)))
   (nth 2)
   (/ 360)
   (double)
   (math/abs)))

(defn get-days-in-year
  "Given a year return the number of days in the year"
  [y]
  (if (= 0 (mod y 4))
    (if (= 0 (mod y 100))
      (if (= 0 (mod y 400))
        366.0
        365.0)
      366.0)
    365.0))

(defn act-act-diff
  "Calculate the fractional years between two excel serial dates using the actual
  number of days between them and using a denominator of the actual days in each 
  year"
  [date-1 date-2]
  (let [[y1] (-> (min date-1 date-2)
                 (excel/build-calendar)
                 (excel/extract-date-fields))
        [y2] (-> (max date-1 date-2)
                 (excel/build-calendar)
                 (excel/extract-date-fields))
        [year-count total-days]
        (reduce (fn [[c1 c2] c-y]
                  [(inc c1)
                   (+ c2
                      (get-days-in-year c-y))])
                [0.0 0.0]
                (range y1 (inc y2)))]
    (/ (- date-2 date-1)
       (/ total-days year-count))))

(defn get-cell-type
  "Get the type of a cell as either :unknown :string or :boolean"
  [c]
  (cond
    (= CellType/FORMULA (.getCellType c))
    (cond
      (= CellType/NUMERIC (.getCachedFormulaResultType c)) :numeric
      (= CellType/STRING (.getCachedFormulaResultType c)) :string
      (= CellType/BOOLEAN (.getCachedFormulaResultType c)) :boolean
      (= CellType/ERROR (.getCachedFormulaResultType c)) :error
      :else :unknown)
    (= CellType/NUMERIC (.getCellType c)) :numeric
    (= CellType/STRING (.getCellType c)) :string
    (= CellType/BOOLEAN (.getCellType c)) :boolean
    (= CellType/BLANK (.getCellType c)) :empty
    (= CellType/ERROR (.getCellType c)) :error
    :else
    :unknown))

(defn extract-test-formulas
  "Extract some test formulas and results from an Excel Workbook.
   Assume that the ones we want are in the second column of the
   sheet in the workbook (no headers)"
  [wb-name sheet-name]
  (->> (dk/load-workbook-from-resource wb-name)
       (dk/select-sheet sheet-name)
       dk/row-seq
       (map dk/cell-seq)
       (reduce
        (fn [accum [_ c2]]
          (when (= CellType/FORMULA (.getCellType c2))
            (let [value-type (get-cell-type c2)
                  cell-style (.getCellStyle c2)
                  cell-address (.getAddress c2)]
              (conj accum
                    (let [cell-value (case value-type
                                       :numeric (.getNumericCellValue c2)
                                       :string (.getStringCellValue c2)
                                       :boolean (.getBooleanCellValue c2)
                                       :empty nil
                                       :error "#ERROR"
                                       "")
                          look-like-date? (and (= :numeric value-type)
                                               (DateUtil/isCellDateFormatted c2))]
                      (cond->
                       {:type value-type
                        :formula (.getCellFormula c2)
                        :format (.getDataFormatString cell-style)
                        :address (.formatAsString cell-address)
                        :row (.getRow cell-address)
                        :column (.getRow cell-address)
                        :value cell-value}
                        look-like-date?
                        (merge
                         {:excel-date-value (.getDateCellValue c2)
                          :calc-date-value (excel-serial-date->local-date-time cell-value)})))))))
        [])))

(comment
  (extract-test-formulas "TEST1.xlsx" "Sheet1")
  :end)