(ns ^{:author "Yuji Hamaguchi"}
  excelAsRDS.Dml
  (:require
    [clojure.data.json :as json :only [read-str read-json write-str]]
    [clojure.set :as set :only [difference]]
    [clojure.java.io :as io :only [reader]])
  (:import
    (org.apache.poi.ss.usermodel Row Cell DateUtil)
    (org.apache.poi.ss.usermodel Workbook WorkbookFactory Cell DataFormatter FormulaEvaluator)
    (java.io File FileInputStream FileOutputStream))
  (:gen-class
    :name excelAsRDS.Dml
    :methods [
      #^{:static true} [getSSCellValues [String Integer String] String]
      #^{:static true} [setSSCellValues [String Integer String] void]
      #^{:static true} [selectSS [String String String] String]
      #^{:static true} [updateSS [String String String] void]
      #^{:static true} [insertSS [String String String] void]]))



(defn get-cell-value
  "Return a cell value."
  [sheet col-idx row-idx]
  (let [row (.getRow sheet row-idx)]
    (if (nil? row)
      ""
      (try
        (let [
          cell (.getCell row col-idx (. Row CREATE_NULL_AS_BLANK))
          wb (.getWorkbook sheet)
          ev (.createFormulaEvaluator (.getCreationHelper wb))
          cell-value (.evaluate ev cell)
          cell-type (if cell-value (.getCellType cell-value) (.getCellType cell))
          ]
          (cond
            (= cell-type Cell/CELL_TYPE_BOOLEAN) (.getBooleanValue cell-value)
            (= cell-type Cell/CELL_TYPE_NUMERIC)
              (if (DateUtil/isCellDateFormatted cell)
                (.formatRawCellContents (DataFormatter.) (.getNumericCellValue cell) -1 "yyyy/mm/dd")
                (.getNumberValue cell-value))
            (= cell-type Cell/CELL_TYPE_STRING) (.getStringValue cell-value)
            (= cell-type Cell/CELL_TYPE_BLANK) ""
            (= cell-type Cell/CELL_TYPE_ERROR) ""
            :else ""))
        ; Out of range for column returns blank value.
        (catch IllegalArgumentException e
          (if (re-find #"^Invalid column index" (.getMessage e))
            ""
            (throw (IllegalArgumentException. (.getMessage e)))))))))

(defn set-cell-value
  "Set a cell to a value."
  [sheet col-idx row-idx val]
  (let [row (.getRow sheet row-idx)]
    (if (nil? row)
      ""
      (try
        (let [
          cell (.getCell row col-idx (. Row CREATE_NULL_AS_BLANK))]
          (if (integer? val)
            (.setCellValue cell (double val))
            (.setCellValue cell val)))
          ; Out of range for column returns blank value.
          (catch IllegalArgumentException e
            (if (re-find #"^Invalid column index" (.getMessage e))
              ""
              (throw (IllegalArgumentException. (.getMessage e)))))))))

(defn set-cell-formula
  "Set a cell to a excel formula."
  [sheet col-idx row-idx formula]
  (let [row (.getRow sheet row-idx)]
    (if (nil? row)
      ""
      (try
        (let [
          cell (.getCell row col-idx (. Row CREATE_NULL_AS_BLANK))]
          (.setCellFormula cell formula))
          ; Out of range for column returns blank value.
          (catch IllegalArgumentException e
            (if (re-find #"^Invalid column index" (.getMessage e))
              ""
              (throw (IllegalArgumentException. (.getMessage e)))))))))


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
      (catch Exception e false))))

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
      (catch Exception e false))))

(defn meet-where-clause-cond
  "Returns true if a row meets conditions in the WHERE clause, false otherwise."
  [{col-idx-map :columnIndex} sheet row-idx where-clause]
  (let [cond-keys (keys where-clause)]
    (letfn [(meet-cond [cond-keys]
      (if (empty? cond-keys)
        true
        (and
          (let [
            key (first cond-keys)
            ex-val (key where-clause)
            val (get-cell-value sheet (key col-idx-map) row-idx)]
            (if (string? val)
              (= ex-val val)
              (== ex-val val)))
          (meet-cond (rest cond-keys)))))]
    (meet-cond cond-keys))))

(defn exist-required-value
  "Returns true if a row contains required attributes, false otherwise."
  [{col-idx-map :columnIndex required :required} sheet row-idx]
  (not-any?
    (fn [col-idx] (empty? (get-cell-value sheet col-idx row-idx)))
    (map second (filter (fn [col] ((set (map keyword required)) (first col))) col-idx-map))))

(defn load-schema-info [schema-file-name]
  "Load schema definition."
  (let [
    required-attrs
    #{:sheetIndex
      :columnIndex
      :startRowIndex
      :endRowIndex
      :required}
      schema-info (with-open [rdr (io/reader schema-file-name)] (json/read-json (apply str (line-seq rdr))))
      diff (set/difference required-attrs (set (keys schema-info)))]
      (if (not (empty? diff))
        (throw (Exception. (str "Required attributes (" diff ") not exist in file (" schema-file-name ").")))
        (cond
          (= 0 (count (schema-info :columnIndex)))
          (throw (Exception. (str "Column index definition (key 'columnIndex') not exist in file (" schema-file-name ").")))
          :else schema-info))))

(defn selectSS
  "Returns JSON string that map collection is selected from excel spreadsheet."
  [schema-file-name xls-file-name select-stmt-json]
  (let [
    schema-info (load-schema-info schema-file-name)
    select-stmt (json/read-json select-stmt-json)
    attrs (let [attrs (:attributes select-stmt)] (if (empty? attrs) (map name (keys (schema-info :columnIndex))) attrs))
    where-clause (:whereClause select-stmt)]
    ; Attribute in select clause does not exists.
    (let [diff (set/difference (set attrs) (set (map name (keys (schema-info :columnIndex)))))]
      (if (not (empty? diff)) (throw (RuntimeException. (str "Attributes (" diff ") not exist in select statement.")))))
    ; Attribute in where clause does not exists.
    (let [diff (set/difference (set (map name (keys where-clause))) (set (map name (keys (schema-info :columnIndex)))))]
      (if (not (empty? diff)) (throw (RuntimeException. (str "Attributes (" diff ") not exist in where clause.")))))
    (json/write-str
      (if (empty? attrs)
        []
        (with-open [in (FileInputStream. xls-file-name)]
          (let [
            workbook (WorkbookFactory/create in)
            sheet (.getSheetAt workbook (schema-info :sheetIndex))]
            (set
              (map
                #(apply hash-map
                  (mapcat
                    (fn [attr]
                      [attr (get-cell-value sheet ((schema-info :columnIndex) (keyword attr)) %)])
                    attrs))
                (filter
                  #(and
                    (exist-required-value schema-info sheet %)
                    (meet-where-clause-cond schema-info sheet % where-clause))
                  (range (schema-info :startRowIndex) (+ 1 (schema-info :endRowIndex))))))))))))

(defn -selectSS [schema-file-name xls-file-name attrs]
  (selectSS schema-file-name xls-file-name attrs))

(defn getSSCellValues
  "Return JSON string that value collection is got from excel spreadsheet."
  [xls-file-name sheet-idx addrs]
  (if (not (is-valid-cell-addr-coll addrs))
    (throw (RuntimeException. (str "Invalid cell address list. (" addrs ")")))
    (let [addrs (json/read-json addrs)]
      (json/write-str
        (with-open [in (FileInputStream. xls-file-name)]
          (let [
            workbook (WorkbookFactory/create in)
            sheet (.getSheetAt workbook sheet-idx)]
            (map (fn [addr] (get-cell-value sheet (first addr) (second addr))) addrs)))))))

(defn -getSSCellValues [xls-file-name sheet-idx addrs]
  (getSSCellValues xls-file-name sheet-idx addrs))

(defn setSSCellValues
  "Set values to excel spreadsheet."
  [xls-file-name sheet-idx key-value-coll-json]
  (if (not (is-valid-cell-addr-val-coll key-value-coll-json))
    (throw (RuntimeException. (str "Invalid cell address and value list. (" key-value-coll-json ")")))
    (let [
      key-value-coll (json/read-json key-value-coll-json)
      in (FileInputStream. xls-file-name)]
      (try
        (let [
          workbook (WorkbookFactory/create in)
          sheet (.getSheetAt workbook sheet-idx)]
          (doseq [kv key-value-coll] (set-cell-value sheet (nth kv 0) (nth kv 1) (nth kv 2)))
          (with-open [out (FileOutputStream. xls-file-name)]
            (.write workbook out)))
        (finally
          (.close in))))))

(defn -setSSCellValues [xls-file-name sheet-idx key-value-coll-json]
  (setSSCellValues xls-file-name sheet-idx key-value-coll-json))

(defn insertSS
  "Insert data to excel spreadsheet."
  [schema-file-name xls-file-name key-value-map-set-json]
  (let [
    schema-info (load-schema-info schema-file-name)
    {
      col-idx-map :columnIndex
      stt-row-idx :startRowIndex
      end-row-idx :endRowIndex
      sheet-idx :sheetIndex
      required :required
    } schema-info
    kvms (json/read-json key-value-map-set-json)
    in (FileInputStream. xls-file-name)]
    (try
      (let [
        workbook (WorkbookFactory/create in)
        sheet (.getSheetAt workbook sheet-idx)]
        (letfn [
          ; Generate cell addresses for values.
          (gen-addr-val-map [kvm row-idx]
            (if (empty? kvm)
              ()
              (let
                [kv (first kvm)
                col-idx (col-idx-map (key kv))
                val (val kv)] 
                (cons (vector col-idx row-idx val) (gen-addr-val-map (rest kvm) row-idx)))))]
          (let [
            valid-kvms (let [
              req-set (set (map keyword required))
              col-set (set (keys col-idx-map))]
              (filter
                (complement (fn [kvm]
                  (let [usr-set (set (keys kvm))]
                  ; Required attribute is missing or non-existent attribute is provided.
                  (if
                    (not
                      (and
                        (empty? (set/difference req-set usr-set))
                        (empty? (set/difference usr-set col-set))))
                    (throw
                      (RuntimeException.
                        (str "Record (" kvm ") is not consistent with schema definition in the file (" schema-file-name ").")))))))
                kvms))
            avl-row-idxs (filter
              (complement (partial exist-required-value schema-info sheet))
              (range stt-row-idx (inc end-row-idx)))]
            ; Can't get available row.
            (if (> (count valid-kvms) (count avl-row-idxs))
              (throw
                (RuntimeException.
                  (str (str (count valid-kvms)) " rows insert failed (all row). Available row count is " (str (count avl-row-idxs)) "."))))                
            (doseq [kv (mapcat #(gen-addr-val-map %1 %2) valid-kvms avl-row-idxs)]
              (set-cell-value sheet (nth kv 0) (nth kv 1) (nth kv 2)))))
        (with-open [out (FileOutputStream. xls-file-name)]
          (.write workbook out)))
      (finally
        (.close in)))))

(defn -insertSS [schema-file-name xls-file-name key-value-map-set-json]
  (insertSS schema-file-name xls-file-name key-value-map-set-json))

(defn updateSS
  "Update data in excel spreadsheet."
  [schema-file-name xls-file-name update-stmts-json]
  (let [
    schema-info (load-schema-info schema-file-name)
    {
      col-idx-map :columnIndex
      stt-row-idx :startRowIndex
      end-row-idx :endRowIndex
      sheet-idx :sheetIndex
      required :required
    } schema-info
    update-stmts (json/read-json update-stmts-json)
    in (FileInputStream. xls-file-name)]
    (try
      (let [
        workbook (WorkbookFactory/create in)
        sheet (.getSheetAt workbook sheet-idx)]
        (letfn [
          (gen-addr-val-map-from-upd-stmt [upd-stmt]
            (let [
              kvm (dissoc upd-stmt :whereClause)
              where-clause (upd-stmt :whereClause)]
              (letfn [(cnv-kv [kvm row-idxs]
                (if (empty? kvm)
                  ()
                  (let
                    [kv (first kvm)
                    col-idx (col-idx-map (key kv))
                    val (val kv)]
                    (concat (for [row-idx row-idxs] (vector col-idx row-idx val)) (cnv-kv (rest kvm) row-idxs)))))]
                (let [row-range (range stt-row-idx (inc end-row-idx)) ]
                  (if (empty? where-clause)
                    (cnv-kv kvm row-range)
                    (let [meet-rows (filter #(meet-where-clause-cond schema-info sheet % where-clause) row-range)]
                      (if (not (empty? meet-rows))
                        (cnv-kv kvm meet-rows))))))))]
          (let [
            valid-upd-stmts
              (let [
                req-set (set (map keyword required))
                col-set (set (keys col-idx-map))]
                (filter
                  (complement (fn [kvm]
                    (let [
                      usr-map (dissoc kvm :whereClause)
                      usr-set (set (keys usr-map))
                      usr-null-set (set (keys (filter #(empty? (val %)) usr-map)))
                      usr-where-set (set (keys (kvm :whereClause)))]
                      (if
                        (not
                          (and
                            (empty? (set/difference usr-set col-set))
                            (= req-set (set/difference req-set usr-null-set))
                            (empty? (set/difference usr-where-set col-set))))
                        (throw
                          (RuntimeException.
                            (str "Record (" kvm ") is not consistent with schema definition in the file (" schema-file-name ").")))))))
                  update-stmts))]
            (doseq [kv (mapcat #(gen-addr-val-map-from-upd-stmt %1) valid-upd-stmts)]
              (set-cell-value sheet (nth kv 0) (nth kv 1) (nth kv 2)))))
        (with-open [out (FileOutputStream. xls-file-name)]
          (.write workbook out)))
      (finally
        (.close in)))))

(defn -updateSS [schema-file-name xls-file-name update-stmts-json]
  (updateSS schema-file-name xls-file-name update-stmts-json))
