(ns scratch
  (:require [graph :refer [eval-range] :as graph]
            [excel :as excel]
            [ast-processing :as ast]
            [xlparse :as parser]
            [report :as report]
            [xlparse :as parse]
            [shunting :as sh]
            [functions])
  (:import [java.util Locale TimeZone Calendar GregorianCalendar Calendar$Builder]
           [java.time LocalDate]
           [org.apache.poi.util LocaleUtil]
           [org.apache.poi.ss SpreadsheetVersion]
           
           [org.apache.poi.ss.usermodel BuiltinFormats DateUtil CellType]
           [org.apache.poi.ss.formula.atp DateParser]
           [java.text DateFormatSymbols]))

(comment
  (graph/explain-workbook "TEST1.xlsx" "Sheet2")

  (def WB-MAP
    (-> "TEST1.xlsx"
        (graph/explain-workbook "Sheet2")
        (graph/get-cell-dependencies)
        (graph/add-graph)
        (graph/connect-disconnected-regions)))

  (graph/expand-cell-range "Sheet2!$L$4:$N$6" WB-MAP)
  (meta (graph/expand-cell-range "Sheet2!$L$4:$N$6" WB-MAP))

  (graph/eval-range "Sheet2!$L$4:$N$6" WB-MAP)
  (meta (graph/eval-range "Sheet2!$L$4:$N$6" WB-MAP))

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
  
  (report/get-formulas-by-usage "TEST-cyclic.xlsx")
  :end)

(comment
  (require '[functions :as f])
  (require '[ubergraph.core :as uber])
  {:vlaaad.reveal/command '(clear-output)}

  (clojure.string/index-of "database" "base")
  (f/fn-search "base" "database")

  (def WB-MAP
    (-> "TEST-cyclic.xlsx"
        (graph/explain-workbook)
        (graph/get-cell-dependencies)
        (graph/add-graph)
        (graph/connect-disconnected-regions)))

  (graph/expand-cell-range "Sheet1!B1:B3" WB-MAP)
  (graph/expand-cell-range "Sheet3!B1:B3" WB-MAP)

  (graph/eval-range "Sheet3!B1:B3" WB-MAP)

  (functions/fn-search (str "B") (str "ABC") 1.0)

  (-> "=SUM(OFFSET(A1,1,1):C3)"
      (parse/parse-to-tokens-pass-1)
      (parse/parse-to-tokens-pass-2)
      (parse/parse-to-tokens-pass-3)
      (parse/parse-to-tokens-pass-4))

  (def P3 (-> "=SUM(1:1)"
              (parse/parse-to-tokens-pass-1)
              (parse/parse-to-tokens-pass-2)
              (parse/parse-to-tokens-pass-3)
              (parse/parse-to-tokens-pass-4)))


  (reduce
   (fn [accum
        [{prev-type :type prev-sub-type :sub-type :as prev-token}
         {current-type :type current-sub-type :sub-type current-value :value :as current-token}
         {next-type :type next-sub-type :sub-type next-value :value :as next-token}]]
     (cond
       (nil? current-token)
       accum
       (and (contains? #{:Range :Number} prev-sub-type)
            (= current-type :OperatorInfix)
            (= ":" current-value)
            (contains? #{:Range :Number} next-sub-type))
       (conj  accum current-token)
       :else
       (conj  accum current-token)))
   []
   (partition 3 1 (into [] (concat [nil] P3 [nil]))))

  (loop [part-tokens (partition 3 1 (into [] (concat [nil] P3 [nil]))) r []]
    (if-not (seq part-tokens)
      r
      (let [[{prev-type :type prev-sub-type :sub-type prev-value :value :as prev-token}
             {current-type :type current-sub-type :sub-type current-value :value :as current-token}
             {next-type :type next-sub-type :sub-type next-value :value :as next-token}]
            (first part-tokens)
            consolidate? (and (contains? #{:Range :Number} prev-sub-type)
                              (= current-type :OperatorInfix)
                              (= ":" current-value)
                              (contains? #{:Range :Number} next-sub-type))]
        (recur (cond-> part-tokens
                 true
                 rest
                 consolidate?
                 rest)
               (cond
                 (nil? current-token)
                 r
                 consolidate?
                 (conj (into [] (butlast r)) (assoc current-token
                                                    :type :Operand
                                                    :sub-type :Range
                                                    :value (str prev-value current-value next-value)))
                 :else
                 (conj r current-token))))))

  (reduce
   (fn [accum
        [{prev-type :type prev-sub-type :sub-type :as prev-token}
         {current-type :type current-sub-type :sub-type current-value :value :as current-token}
         {next-type :type next-sub-type :sub-type next-value :value :as next-token}]]
     (cond
       (nil? current-token)
       accum
       (and (contains? #{:Range :Number} prev-sub-type)
            (= current-type :OperatorInfix)
            (= ":" current-value)
            (contains? #{:Range :Number} next-sub-type))
       (conj  accum current-token)
       :else
       (conj  accum current-token)))
   []
   (partition 3 1 (into [] (concat [nil] P3 [nil]))))

  (partition 3 1 (into [] (concat [nil] P3 [nil])))

  (-> "=INDEX(A2:C6, 5, 2)"
      (parse/parse-to-tokens)
      #_(parse/nest-ast)
      #_(parse/wrap-ast)
      #_(ast/process-tree)
      #_(sh/parse-expression-tokens)
      #_(ast/unroll-for-code-form "Sheet1"))

  (-> "=SUM($B$2:X1)"
      (parse/parse-to-tokens)
      #_(parse/nest-ast)
      #_(parse/wrap-ast)
      #_(ast/process-tree)
      #_(sh/parse-expression-tokens)
      #_(ast/unroll-for-code-form "Sheet5"))

  (-> "=SUM($B$2:X1)"
      (parse/parse-to-tokens)
      #_(parse/nest-ast)
      #_(parse/wrap-ast)
      #_(ast/process-tree)
      #_(sh/parse-expression-tokens)
      #_(ast/unroll-for-code-form "Sheet5"))

  (->
   #_"=INDEX(A2:C6, 2, 3)"
   #_"=INDEX((A1:C6, A8:C11), 2, 2, 2)"
   #_"=SUM(INDEX(A1:C11, 0, 3, 1))"
   #_"=SUM($B$2:INDEX(A2:C6, 5, 2))"
   #_"=INDEX(A2:C6, 5, 2)"
   #_"=SUM(B2:INDEX(B2:B3,2))"
   "=INDIRECT(A12)"
   (parse/parse-to-tokens)
   (parse/nest-ast)
   (parse/wrap-ast)
   (ast/process-tree)
   (sh/parse-expression-tokens)
   (ast/unroll-for-code-form "Sheet5"))

  (binding [graph/*context* WB-MAP]
    (-> '(functions/fn-sum (functions/fn-range (eval-range "Sheet5!B2") (functions/fn-index (eval-range "Sheet5!B2:B3") 2.0)))
        (graph/substitute-ranges)
        #_(eval)))

  (eval-range "Sheet5!A1:C11" WB-MAP)
  (meta (eval-range "Sheet5!A1:C11" WB-MAP))
  (functions/fn-sum (functions/fn-index (#'graph/eval-range "Sheet5!A1:C11" WB-MAP) 0.0 3.0 1.0))

  (functions/fn-index (graph/eval-range "Sheet5!A2:C6" WB-MAP) 2.0 3.0)

  (functions/fn-sum
   (functions/fn-range (eval-range "Sheet5!$B$2" WB-MAP) 
                       (functions/fn-index (eval-range "Sheet5!A2:C6" WB-MAP) 5.0 2.0)))
  
  (functions/fn-sum
   (functions/fn-range (eval-range "Sheet5!$B$2" WB-MAP)
                       (functions/fn-index-reference (eval-range "Sheet5!A2:C6" WB-MAP) 5.0 2.0)))

  (require '[clojure.walk :as walk])

  graph/eval-range
  (eval (quote eval-range))
  (eval (quote graph/eval-range))

  

  ; eval (functions/fn-index-reference (eval-range "Sheet5!A2:C6" WB-MAP) 5.0 2.0) to get "B6"
  (defn construct-dynamic-range [forms]
    (letfn [(->fn [f]
                  (if (var? (eval f)) 
                    (-> f eval deref) 
                    (eval f)))]
      (let [fs (mapv
                (fn [[f-name f-arg :as form]]
                  (cond (= functions/fn-index (->fn f-name))
                        (re-matches #"(.*!)?(.*)" (eval (cons 'functions/fn-index-reference (rest form))))
                        (= graph/eval-range (->fn f-name))
                        (re-matches #"(.*!)?(.*)" f-arg)
                        :else
                        (throw (IllegalArgumentException.
                                (str "NO MATCH. "
                                     "Form " (pr-str form)
                                     " FNAME:" f-name
                                     " M?:" (if (var? (eval f-name))
                                              (-> f-name eval deref)
                                              (eval f-name)))))))
                forms)
          fstr (str (some (fn [[_ sheet _]] (when (some? sheet) sheet)) fs)
                    (clojure.string/join ":" (map (fn [[_ _ label]] label) fs)))]
    `(eval-range ~fstr WB-MAP))))
  
  (defn construct-dynamic-range-pre [forms]
    (letfn [(->fn [f]
              (if (var? (eval f))
                (-> f eval deref)
                (eval f)))]
      (let [fs (mapv
                (fn [[f-name f-arg :as form]]
                  (cond (= functions/fn-index (->fn f-name))
                        (re-matches #"(.*!)?(.*)" (eval (cons 'functions/fn-index-reference (rest form))))
                        (= graph/eval-range (->fn f-name))
                        (re-matches #"(.*!)?(.*)" f-arg)
                        :else
                        (throw (IllegalArgumentException.
                                (str "NO MATCH. "
                                     "Form " (pr-str form)
                                     " FNAME:" f-name
                                     " M?:" (if (var? (eval f-name))
                                              (-> f-name eval deref)
                                              (eval f-name)))))))
                forms)
            fstr (str (some (fn [[_ sheet _]] (when (some? sheet) sheet)) fs)
                      (clojure.string/join ":" (map (fn [[_ _ label]] label) fs)))]
        `(eval-range ~fstr WB-MAP))))
  
  (construct-dynamic-range 
   '((#'graph/eval-range "Sheet5!$B$2" WB-MAP)
     (functions/fn-index (#'graph/eval-range "Sheet5!A2:C6" WB-MAP) 5.0 2.0)))
  
  (-> graph/eval-range var deref)
  (-> graph/eval-range var)
  (construct-dynamic-range
   '((functions/fn-index (eval-range "Sheet5!A2:C6" WB-MAP) 5.0 2.0)
   (eval-range "Sheet5!$B$2" WB-MAP)))
  
  (var graph/eval-range)
  (var graph/eval-range)

  (binding [graph/*context* WB-MAP]
    (->> '(functions/fn-sum
          (functions/fn-range
           (eval-range "Sheet5!$B$2")
           (functions/fn-index (eval-range "Sheet5!A2:C6") 5.0 2.0)))
         (graph/substitute-ranges)
         #_(walk/postwalk
          (fn [f]
            (if (and (list? f) (= 'functions/fn-range (first f)))
              (construct-dynamic-range (-> f rest))
              f)))
         #_(eval)))
  
  (binding [graph/*context* WB-MAP]
    (-> '(functions/fn-range 
        ((var graph/eval-range) "Sheet5!B2" graph/*context*) 
        (functions/fn-index ((var graph/eval-range) "Sheet5!B2:B3" graph/*context*) 2.0))
      (rest)
      (construct-dynamic-range)))
  
  (->> 
   '((#'graph/eval-range "Sheet5!B2" graph/*context*)
   (functions/fn-index (#'graph/eval-range "Sheet5!B2:B3" graph/*context*) 2.0))
   (map first))
  
  (var graph/eval-range)
  (var? 'graph/eval-range)
  (eval 'graph/eval-range)
  (eval #'graph/eval-range)
  (var? #'graph/eval-range)
  (deref #'graph/eval-range)

  (var? functions/fn-index)
  (eval 'functions/fn-index)
  
  (re-matches #"(.*!)?(.*)" "S!B6")
  (meta *1)

  (rest '(1 (2 3)))

  (def WB-MAP
    (-> "TEST-cyclic.xlsx"
        (graph/explain-workbook)
        (graph/get-cell-dependencies)
        (graph/add-graph)
        (graph/connect-disconnected-regions)))

  (graph/recalc-workbook WB-MAP "Sheet6")

  (graph/expand-cell-range "Sheet6!A14.0" WB-MAP)
  (eval-range
   (str "Sheet6!"
        (functions/fn-indirect
         (eval-range
          "Sheet6!A11" WB-MAP)))
   WB-MAP)
  
  ((partial functions/fn-indirect "Sheet6") (eval-range "Sheet6!A13" WB-MAP))

  ((partial
        functions/fn-indirect
        "Sheet6")
       (str
        (str
         "Sheet6!B")
        (eval-range
         "Sheet6!A14"
         WB-MAP)))
  
  ((partial functions/fn-indirect "Sheet6") (eval-range "Sheet6!A12" WB-MAP))
  
  (eval-range "Sheet6!B11" WB-MAP)
  
  (functions/fn-sum
   (functions/fn-range
    (graph/eval-range
     "Sheet5!$B$2" WB-MAP)
    (functions/fn-index
     (graph/eval-range
      "Sheet5!A2:C6" WB-MAP)
     5.0
     2.0)))
  
  (functions/fn-index 
   (functions/fn-union
    (eval-range "Sheet5!A1:C6" WB-MAP)
    (eval-range "Sheet5!A8:C11" WB-MAP)) 
   2.0 2.0 2.0)
  
  (meta (eval-range "Sheet5!A1:C6" WB-MAP))
  (meta (eval-range "Sheet5!A8:C11" WB-MAP))

  (functions/fn-sum
   (graph/eval-range
    "Sheet5!$B$2:B6" WB-MAP))

  (functions/fn-index-reference
   (graph/eval-range
    "Sheet5!A2:C6" WB-MAP)
   5.0
   2.0)


  (defmacro fn-range [[fn-1 f1] [fn-2 f2]]
    f1)

  (macroexpand
   '(fn-range (eval-range "Sheet5!$B$2")
              (functions/fn-index
               (eval-range "Sheet5!A2:C6") 5.0 2.0)))
  (fn-range '(eval-range "Sheet5!$B$2")
            '(functions/fn-index
              (eval-range "Sheet5!A2:C6") 5.0 2.0))


  (CellReference/convertNumToColString 0)

  ; A2 [1 0] row=1 col=0
  ;    [1+5, 0+2]
  (CellReference/convertNumToColString 2)

  (meta (graph/eval-range
         "Sheet5!A2:C6" WB-MAP))

  (some-> '((100.0
             "A")
            (200.0
             "AA")
            (300.0
             "B"))
          (nth  10 nil)
          (nth  1 nil))
  :end
  )