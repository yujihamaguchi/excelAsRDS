(ns ^{:author "Yuji Hamaguchi"}
  excelAsRDS.Dml
  (:require [clojure.data.json :as json :only [read-str read-json write-str]]
            [clojure.set :as set :only [difference]]
            [clojure.java.io :as io :only [reader]]
            [clojure.string :as str :only [replace]])
  (:import [org.apache.poi.ss.usermodel Row Cell DateUtil]
           [org.apache.poi.ss.usermodel Workbook WorkbookFactory Cell DataFormatter FormulaEvaluator]
           [java.io File FileInputStream FileOutputStream])
  (:gen-class
    :name excelAsRDS.Dml
    :methods [#^{:static true} [getSSCellValues [String Integer String] String]
              #^{:static true} [setSSCellValues [String Integer String] void]
              #^{:static true} [selectSS [String String String] String]
              #^{:static true} [updateSS [String String String] void]
              #^{:static true} [insertSS [String String String] void]]))

(declare get-cell-value
         set-cell-value
         set-cell-formula
         is-valid-cell-addr-coll
         is-valid-cell-addr-val-coll
         meet-where-clause-cond
         exist-required-value
         load-schema-info
         selectSS
         getSSCellValues
         setSSCellValues
         insertSS
         updateSS)

(defn get-cell-value
  "Return a cell value."
  ([sheet col-idx row-idx]
    (if-let [row (.getRow sheet row-idx)]
      (try
        (let [cell (.getCell row col-idx Row/CREATE_NULL_AS_BLANK)
              ev (.createFormulaEvaluator (.getCreationHelper (.getWorkbook sheet)))
              cell-value (.evaluate ev cell)
              cell-type (if cell-value (.getCellType cell-value) (.getCellType cell))]
          (condp = cell-type
            Cell/CELL_TYPE_BOOLEAN (.getBooleanValue cell-value)
            Cell/CELL_TYPE_NUMERIC (if (DateUtil/isCellDateFormatted cell)
                                      (.formatRawCellContents (DataFormatter.)
                                                              (.getNumberValue cell-value)
                                                              -1
                                                              "yyyy/mm/dd")
                                      (.getNumberValue cell-value))
            Cell/CELL_TYPE_STRING  (.getStringValue cell-value)
            Cell/CELL_TYPE_BLANK   ""
            Cell/CELL_TYPE_ERROR   ""
            ""))
        ; Out of range for column returns blank value.
        (catch IllegalArgumentException e
          (if-not (re-find #"^Invalid column index" (.getMessage e))
            (throw (IllegalArgumentException. (.getMessage e)))
            "")))
      ""))
  ([sheet col-idx row-idx formula]
    (set-cell-formula sheet
                      col-idx
                      row-idx
                      (str/replace formula
                                   "_ROWIDX_"
                                   (str (inc row-idx))))
    (get-cell-value sheet col-idx row-idx)))

(defn set-cell-value
  "Set a cell to a value."
  [sheet col-idx row-idx val]
  (if-let [row (.getRow sheet row-idx)]
    (try
      (let [cell (.getCell row
                           col-idx
                           Row/CREATE_NULL_AS_BLANK)]
        (if (integer? val)
          (.setCellValue cell (double val))
          (.setCellValue cell val)))
      ; Out of range for column returns blank value.
      (catch IllegalArgumentException e
        (if-not (re-find #"^Invalid column index" (.getMessage e))
          (throw (IllegalArgumentException. (.getMessage e)))
          "")))
    ""))

(defn set-cell-formula
  "Set a cell to a excel formula."
  [sheet col-idx row-idx formula]
  (if-let [row (.getRow sheet row-idx)]
    (try
      (let [cell (.getCell row
                           col-idx
                           Row/CREATE_NULL_AS_BLANK)]
        (.setCellFormula cell formula))
      ; Out of range for column returns blank value.
      (catch IllegalArgumentException e
        (if-not (re-find #"^Invalid column index" (.getMessage e))
          (throw (IllegalArgumentException. (.getMessage e)))
          "")))
    ""))

(defn is-valid-cell-addr-coll
  "Returns true if cell address collection is valid, false otherwise."
  [cell-addr-coll]
  (and
    (not (empty? cell-addr-coll))
    (try
      (let [addrs (json/read-json cell-addr-coll)]
        (and
          (vector? addrs)
          (<= 1 (count addrs))
          (every? vector? addrs)
          (every? #(= 2 (count %)) addrs)
          (every? #(every? integer? %) addrs)))
      (catch Exception _ false))))

(defn is-valid-cell-addr-val-coll
  "Returns true if cell address and value collection is valid, false otherwise."
  [cell-addr-val-coll]
  (and
    (not (empty? cell-addr-val-coll))
    (try
      (let [key-value-coll (json/read-json cell-addr-val-coll)]
        (and
          (vector? key-value-coll)
          (<= 1 (count key-value-coll))
          (every? vector? key-value-coll)
          (every? #(= 3 (count %)) key-value-coll)
          (every? #(every? integer? (take 2%)) key-value-coll)))
      (catch Exception _ false))))

(defn meet-where-clause-cond
  "Returns true if a row meets conditions in the WHERE clause, false otherwise."
  [{col-idx-map :columnIndex} sheet row-idx where-clause]
  (let [cond-keys (keys where-clause)]
    (letfn [(meet-cond [cond-keys]
      (if (empty? cond-keys)
        true
        (and
          (let [key (first cond-keys)
                ex-val (key where-clause)
                val (get-cell-value sheet
                                    (key col-idx-map)
                                    row-idx)]
            (if (string? val)
              (= ex-val val)
              (== ex-val val)))
          (meet-cond (rest cond-keys)))))]
    (meet-cond cond-keys))))

(defn exist-required-value
  "Returns true if a row contains required attributes, false otherwise."
  [{colnm-idx-map :columnIndex required-attrs :required} sheet row-idx]
  (let [req-attr-set (set (map keyword required-attrs))]
    (not-any? (fn [col-idx]
                (empty? (get-cell-value sheet
                                        col-idx
                                        row-idx)))
              (for [m colnm-idx-map :when (req-attr-set (first m))] (second m)))))

(defn load-schema-info
  "Load schema definition."
  [schema-file-name]
  (let [required-attrs #{:sheetIndex
                         :columnIndex
                         :startRowIndex
                         :endRowIndex
                         :required}
        schema-info (json/read-json (slurp schema-file-name))
        diff (set/difference required-attrs
                             (set (keys schema-info)))]
    (cond
      (not (empty? diff)) (throw (Exception. (str "Required attributes (" diff ") not exist in file (" schema-file-name ").")))
      (zero? (count (schema-info :columnIndex))) (throw (Exception. (str "Column index definition (key 'columnIndex') not exist in file (" schema-file-name ").")))
      :else schema-info)))

(defn selectSS
  "Returns JSON string that map collection is selected from excel spreadsheet."
  [schema-file-name xls-file-name select-stmt-json]
  (let [schema-info (load-schema-info schema-file-name)
        select-stmt (json/read-json select-stmt-json)
        where-clause (:whereClause select-stmt)
        all-attrs (map name (keys (schema-info :columnIndex)))
        attrs (let [{attrs :attributes} select-stmt]
                (if-not (seq attrs)
                  all-attrs
                  attrs))]
    ; Attribute in select clause does not exists.
    (let [diff (set/difference (set attrs)
                               (set all-attrs))]
      (if (seq diff)
        (throw (RuntimeException. (str "Attributes (" diff ") not exist in select statement.")))))
    ; Attribute in where clause does not exists.
    (let [diff (set/difference (->> where-clause
                                    keys
                                    (map name) 
                                    set)
                               (->> all-attrs
                                    set))]
      (if (seq diff)
        (->> (str "Attributes (" diff ") not exist in where clause.")
             RuntimeException.
             throw)))
    (json/write-str
      (if-not (seq attrs)
        []
        (with-open [in (FileInputStream. xls-file-name)]
          (let [workbook (WorkbookFactory/create in)
                sheet (.getSheetAt workbook (schema-info :sheetIndex))]
            (set
              (map #(apply hash-map
                           (mapcat (fn [attr]
                                     (let [col-idx ((schema-info :columnIndex) (keyword attr))]
                                       (if (and (schema-info :excelFormula)
                                                ((schema-info :excelFormula) (keyword attr)))
                                           [attr (get-cell-value sheet
                                                                 col-idx
                                                                 %
                                                                 ((schema-info :excelFormula) (keyword attr)))]
                                           [attr (get-cell-value sheet
                                                                 col-idx
                                                                 %)])))
                                   attrs))
                   (filter #(and (exist-required-value schema-info
                                                       sheet
                                                       %)
                                 (meet-where-clause-cond schema-info
                                                         sheet
                                                         %
                                                         where-clause))
                     (range (schema-info :startRowIndex)
                            (inc (schema-info :endRowIndex))))))))))))

(defn -selectSS [schema-file-name xls-file-name attrs]
  (selectSS schema-file-name
            xls-file-name
            attrs))

(defn getSSCellValues
  "Return JSON string that value collection is got from excel spreadsheet."
  [xls-file-name sheet-idx addrs]
  (if-not (is-valid-cell-addr-coll addrs)
    (throw (RuntimeException. (str "Invalid cell address list. (" addrs ")")))
    (let [addrs (json/read-json addrs)]
      (json/write-str
        (with-open [in (FileInputStream. xls-file-name)]
          (let [workbook (WorkbookFactory/create in)
                sheet (.getSheetAt workbook sheet-idx)]
            (map (fn [addr]
                   (get-cell-value sheet
                                   (first addr)
                                   (second addr)))
                 addrs)))))))

(defn -getSSCellValues [xls-file-name sheet-idx addrs]
  (getSSCellValues xls-file-name
                   sheet-idx
                   addrs))

(defn setSSCellValues
  "Set values to excel spreadsheet."
  [xls-file-name sheet-idx key-value-coll-json]
  (if-not (is-valid-cell-addr-val-coll key-value-coll-json)
    (throw (RuntimeException. (str "Invalid cell address and value list. (" key-value-coll-json ")")))
    (let [key-value-coll (json/read-json key-value-coll-json)
          in (FileInputStream. xls-file-name)]
      (try
        (let [workbook (WorkbookFactory/create in)
              sheet (.getSheetAt workbook sheet-idx)]
          (doseq [kv key-value-coll]
            (set-cell-value sheet
                            (nth kv 0)
                            (nth kv 1)
                            (nth kv 2)))
          (with-open [out (FileOutputStream. xls-file-name)]
            (.write workbook out)))
        (finally
          (.close in))))))

(defn -setSSCellValues [xls-file-name sheet-idx key-value-coll-json]
  (setSSCellValues xls-file-name
                   sheet-idx
                   key-value-coll-json))

(defn insertSS
  "Insert datas in tuples to excel records."
  [schema-file-name xls-file-name key-value-map-set-json]
  (let [schema-info (load-schema-info schema-file-name)
        { col-idx-map :columnIndex
          stt-row-idx :startRowIndex
          end-row-idx :endRowIndex
          sheet-idx :sheetIndex
          required :required } schema-info
        kvms (json/read-json key-value-map-set-json)
        in (FileInputStream. xls-file-name)]
    (try
      (let [workbook (WorkbookFactory/create in)
            sheet (.getSheetAt workbook sheet-idx)]
        ; Generate cell addresses for values.
        (letfn [(gen-addr-val-map [kvm row-idx]
                  (when (seq kvm)
                    (let [kv (first kvm)
                          col-idx (col-idx-map (key kv))
                          val (val kv)]
                      (cons (vector col-idx row-idx val)
                            (gen-addr-val-map (rest kvm)
                                              row-idx)))))]
          (let [valid-kvms (let [req-set (set (map keyword required))
                                 col-set (set (keys col-idx-map))]
                                  (filter (complement
                                            (fn [kvm]
                                              (let [upd-k-set (set (keys kvm))]
                                                ; Required attribute is missing or non-existent attribute is provided.
                                                (when-not (and (empty? (set/difference req-set upd-k-set))
                                                               (empty? (set/difference upd-k-set col-set)))
                                                  (throw (RuntimeException. (str "Record (" kvm ") is not consistent with schema definition in the file (" schema-file-name ").")))))))
                                     kvms))
                avl-row-idxs (filter (complement (partial exist-required-value
                                                          schema-info
                                                          sheet))
                                     (range stt-row-idx
                                            (inc end-row-idx)))]
            ; Can't get available row.
            (when (> (count valid-kvms) (count avl-row-idxs))
              (throw (RuntimeException. (str (str (count valid-kvms)) " rows insert failed (all row). Available row count is " (str (count avl-row-idxs)) "."))))
            (doseq [kv (mapcat #(gen-addr-val-map %1 %2)
                               valid-kvms
                               avl-row-idxs)]
              (set-cell-value sheet
                              (nth kv 0)
                              (nth kv 1)
                              (nth kv 2)))))
        (with-open [out (FileOutputStream. xls-file-name)]
          (.write workbook out)))
      (finally
        (.close in)))))

(defn -insertSS [schema-file-name xls-file-name key-value-map-set-json]
  (insertSS schema-file-name
            xls-file-name
            key-value-map-set-json))

(defn updateSS
  "Update selected datas in Cells."
  [schema-file-name xls-file-name update-stmts-json]
  (let [schema-info (load-schema-info schema-file-name)
        { col-idx-map :columnIndex
          stt-row-idx :startRowIndex
          end-row-idx :endRowIndex
          sheet-idx :sheetIndex
          required :required } schema-info
        update-stmts (json/read-json update-stmts-json)
        in (FileInputStream. xls-file-name)]
    (try
      (let [workbook (WorkbookFactory/create in)
            sheet (.getSheetAt workbook sheet-idx)]
        (letfn [(gen-addr-val-map-from-upd-stmt [upd-stmt]
                  (let [kvm (dissoc upd-stmt :whereClause)
                        where-clause (upd-stmt :whereClause)]
                    (letfn [(cnv-kv [kvm row-idxs]
                              (when (seq kvm)
                                (let [kv (first kvm)
                                      col-idx (col-idx-map (key kv))
                                      val (val kv)]
                                  (concat (for [row-idx row-idxs] (vector col-idx row-idx val))
                                          (cnv-kv (rest kvm) row-idxs)))))]
                      (let [row-range (range stt-row-idx (inc end-row-idx))]
                        (if-not (seq where-clause)
                          (cnv-kv kvm row-range)
                          (let [meet-rows (filter #(meet-where-clause-cond schema-info
                                                                           sheet
                                                                           %
                                                                           where-clause)
                                                  row-range)]
                            (when (seq meet-rows)
                              (cnv-kv kvm meet-rows))))))))]
          (let [valid-upd-stmts (let [req-set (set (map keyword required))
                                      col-set (set (keys col-idx-map))]
                                  (filter (complement (fn [kvm]
                                                        (let [upd-kv-map (dissoc kvm :whereClause)
                                                              upd-k-set (set (keys upd-kv-map))
                                                              upd-emp-v-k-set (set (keys (filter #(empty? (val %)) upd-kv-map)))
                                                              upd-where-set (set (keys (kvm :whereClause)))]
                                                          (if-not (and (empty? (set/difference upd-k-set col-set))
                                                                       (= req-set (set/difference req-set upd-emp-v-k-set))
                                                                       (empty? (set/difference upd-where-set col-set)))
                                                            (throw (RuntimeException. (str "Record (" kvm ") is not consistent with schema definition in the file (" schema-file-name ").")))))))
                                          update-stmts))]
            (doseq [kv (mapcat #(gen-addr-val-map-from-upd-stmt %1)
                               valid-upd-stmts)]
              (set-cell-value sheet
                              (nth kv 0)
                              (nth kv 1)
                              (nth kv 2)))))
        (with-open [out (FileOutputStream. xls-file-name)]
          (.write workbook out)))
      (finally
        (.close in)))))

(defn -updateSS [schema-file-name xls-file-name update-stmts-json]
  (updateSS schema-file-name
            xls-file-name
            update-stmts-json))
