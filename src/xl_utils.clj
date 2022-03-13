(ns xl-utils
  "A set of utility functions to manipulate sheets names, cell names
   etc. used by Excel.
   TODO: move all the \"!\" functions into here from the existing 
   namespaces, where we currently build references using string manipulation 
   function. Do this so we have one place to look."
  (:require
   [box])
  (:import
   [box Box]))

(defn get-complete-ref-str
  "Given a cell reference make it absolute by adding the sheet name,
   unless it is already an absolute reference, in which case return
   the original absolute reference."
  [sheet-name ref-text]
  (let [value (if (instance? Box ref-text) @ref-text ref-text)
        [_ sheet-with-exclam _] (re-matches #"(.*!)?(.*)" value)]
    (if sheet-with-exclam
      value
      (if sheet-name
        (str sheet-name "!" value)
        value))))

(comment
  (get-complete-ref-str "sheetname" "C1")
  (get-complete-ref-str "sheetname" "sheetname2!C1")
  (get-complete-ref-str "sheetname" (Box. "C1" {}))
  (get-complete-ref-str "sheetname" (Box. "sheetname2!C1" {}))
  :end)
