(ns excel
  (:require
   [clojure.string :as str]
   [dk.ative.docjure.spreadsheet :as dk]
   [clojure.math.numeric-tower :as math])
  (:import
   [java.util Locale TimeZone Calendar Calendar$Builder]
   [java.time LocalDate LocalTime LocalDateTime]
   [java.time.format DateTimeFormatter]
   [org.apache.poi.ss.usermodel CellType DateUtil]))

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
   dates and the number of NASD days, as calculated by MASD 30/360, between
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


(defn get-cell-type
  "Get the type of a cell as either :unknown :string or :boolean"
  [c]
  (cond
    (= CellType/FORMULA (.getCellType c))
    (cond
      (= CellType/NUMERIC (.getCachedFormulaResultType c)) :numeric
      (= CellType/STRING (.getCachedFormulaResultType c)) :string
      (= CellType/BOOLEAN (.getCachedFormulaResultType c)) :boolean
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