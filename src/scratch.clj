(ns scratch
  (:require [graph :refer [eval-range] :as graph]
            [excel :as excel]
            [ast-processing :as ast]
            [xlparse :as parser])
  (:import [java.util Locale TimeZone Calendar GregorianCalendar Calendar$Builder]
           [java.time LocalDate]
           [org.apache.poi.util LocaleUtil]
           [org.apache.poi.ss SpreadsheetVersion]
           [org.apache.poi.ss.util AreaReference CellReference]
           [org.apache.poi.ss.usermodel BuiltinFormats DateUtil CellType]
           [org.apache.poi.ss.formula.atp DateParser]
           [java.text DateFormatSymbols]))

(comment
  (graph/explain-workbook "TEST1.xlsx" "Sheet2")

  (def WB-MAP
    (-> "TEST1.xlsx"
        (graph/explain-workbook "Sheet2")
        (graph/get-cell-dependencies)
        (graph/add-graph)))

  (graph/expand-cell-range "Sheet2!$L$4:$N$6" (:named-ranges WB-MAP))
  (meta (graph/expand-cell-range "Sheet2!$L$4:$N$6" (:named-ranges WB-MAP)))

  (graph/eval-range "Sheet2!$L$4:$N$6" WB-MAP)
  (meta (graph/eval-range "Sheet2!$L$4:$N$6" WB-MAP))

  (let [a {:val ["L1" 0.1 0.0 "L2" 0.2 30.0 "L3" 0.3 35.0]
   :meta {:single? false
          :column? false
          :sheet-name "Sheet2"
          :tl-name "L4"
          :tl-coord [3 11]
          :cols 3
          :rows 3}}
        {a-val :val {:keys [single? column? cols rows]} :meta} a]
    (partition cols a-val))

  WB-MAP

  (excel/build-calendar-for-serial-date 43831.0)

  (-> (excel/build-calendar-for-serial-date 31048.0)
      (excel/extract-date-fields))
  
  (-> (excel/build-calendar-for-serial-date 43922.0)
      (excel/extract-date-fields))
  
  (excel/act-act-diff 31048.0 43922.0)
  (excel/act-act-diff 31048.0 31049.0)
  (excel/nasd-360-diff 43831.0 44408.0)
  (excel/euro-360-diff 43831.0 44408.0)

  :end)

(comment
  
  (excel/get-symbol-match-string :short-month)
    
  (excel/parse-excel-string-to-date-info "1/15/2021")
  (excel/parse-excel-string-to-date-info "2021/1/15") 
  (excel/parse-excel-string-to-date-info "2021-01-15")
  (excel/parse-excel-string-to-date-info "January 15, 2021")
  (excel/parse-excel-string-to-date-info "Jan 2021")
  (excel/parse-excel-string-to-date-info "1/15")
  (excel/parse-excel-string-to-date-info "2021-01-02")
  (excel/parse-excel-string-to-date-info "2021-01-02")

  :end)

(comment 
  
  ;; get functions used in workbook in order of frequency
  (->> (excel/extact-all-formulas-from-workbook "TEST1.xlsx")
       (reduce (fn [accum [sheet-name formulas]]
                 (concat accum
                         (mapcat (fn [{:keys [formula sheet-name reference]}]
                                   (->> (str "=" formula)
                                        (parser/parse-to-tokens)
                                        (keep #(when (and (= :Function (:type %))
                                                          (= :Start (:sub-type %)))
                                                 {:formula (:value %)
                                                  :reference (str sheet-name "!" reference)}))))
                                 formulas)))
               [])
       (reduce (fn [accum {:keys [formula reference]}]
                 (update accum
                         formula
                         (fn [{:keys [count references]}]
                           {:count (inc (or count 0))
                            :references (conj (or references []) reference)})))
               {})
       (sort-by (fn [[formula {:keys [count] :as formula-record}]]
                  count))
       (reverse)
       (map first))
  :end)