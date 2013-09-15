(ns ^{:doc "Operate excel sheet as relational data source using Apache POI."
      :author "Yuji Hamaguchi"}
  excelAsRDS.Dml
  (:require
    [clojure.data.json :as json :only [read-str read-json write-str]]
    [clojure.set :as set :only [difference]]
    [clojure.java.io :as io :only [reader]])
  (:import
    (org.apache.poi.ss.usermodel Row Cell DateUtil)
    (org.apache.poi.hssf.usermodel HSSFWorkbook HSSFCell HSSFDataFormatter HSSFFormulaEvaluator)
    (java.io File FileInputStream FileOutputStream))
  (:gen-class
    :name excelAsRDS.Dml
    :methods [
      #^{:static true} [getSSCellValues [String Integer　String] String]
      #^{:static true} [setSSCellValues [String Integer　String] void]
      #^{:static true} [selectSS [String String　String] String]
      #^{:static true} [updateSS [String String　String] void]
      #^{:static true} [insertSS [String String　String] void]]))

(defn get-cell-value [sheet col-idx row-idx]
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
                (.formatRawCellContents (HSSFDataFormatter.) (.getNumericCellValue cell) -1 "yyyy/mm/dd")
                (.getNumberValue cell-value))
            (= cell-type Cell/CELL_TYPE_STRING) (.getStringValue cell-value)
            (= cell-type Cell/CELL_TYPE_BLANK) ""
            (= cell-type Cell/CELL_TYPE_ERROR) ""
            :else ""))
        ; 列がエクセルの範囲外である場合、空の結果とする
        (catch IllegalArgumentException e
          (if (re-find #"^Invalid column index" (.getMessage e))
            ""
            (throw (IllegalArgumentException. (.getMessage e)))))))))

(defn set-cell-value [sheet col-idx row-idx val]
  (let [row (.getRow sheet row-idx)]
    (if (nil? row)
      ""
      (try
        (let [
          cell (.getCell row col-idx (. Row CREATE_NULL_AS_BLANK))]
          (if (integer? val)
            (.setCellValue cell (double val))
            (.setCellValue cell val)))
          ; 列がエクセルの範囲外である場合、空の結果とする
          (catch IllegalArgumentException e
            (if (re-find #"^Invalid column index" (.getMessage e))
              ""
              (throw (IllegalArgumentException. (.getMessage e)))))))))

(defn is-valid-cell-addr-lst [cell-addr-lst]
  (and
    (not (empty? cell-addr-lst))
    (try
      (let [addrs (json/read-json cell-addr-lst)]
        (and
          (vector? addrs)
          (<= 1 (count addrs))
          (every? vector? addrs)
          (every? #(= 2 (count %)) addrs)
          (every? #(every? integer? %) addrs)))
      (catch Exception e false))))

(defn is-valid-cell-addr-val-lst [cell-addr-val-lst]
  (and
    (not (empty? cell-addr-val-lst))
    (try
      (let [kvs (json/read-json cell-addr-val-lst)]
        (and
          (vector? kvs)
          (<= 1 (count kvs))
          (every? vector? kvs)
          (every? #(= 3 (count %)) kvs)
          (every? #(every? integer? (take 2%)) kvs)))
      (catch Exception e false))))

(defn meet-where-clause-cond [{col-idx-map :columnIndex} sheet row-idx where-clause]
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

(defn exist-required-value [{col-idx-map :columnIndex required :required} sheet row-idx]
  (not-any?
    (fn [col-idx] (empty? (get-cell-value sheet col-idx row-idx)))
    (map second (filter (fn [col] ((set (map keyword required)) (first col))) col-idx-map))))

(defn load-schema-info [sch-file]
  (let [
    required-attrs
    #{:sheetIndex
      :columnIndex
      :startRowIndex
      :endRowIndex
      :required}
      schema-info (with-open [rdr (io/reader sch-file)] (json/read-json (apply str (line-seq rdr))))
      diff (set/difference required-attrs (set (keys schema-info)))]
      (if (not (empty? diff))
        (throw (Exception. (str "Required attributes (" diff ") not exist in file (" sch-file ").")))
        (cond
          (= 0 (count (schema-info :columnIndex)))
          (throw (Exception. (str "Column index definition (key 'columnIndex') not exist in file (" sch-file ").")))
          :else schema-info))))

(defn selectSS
  "select data in excel file."
  {:static true}
  [sch-file xls-file select-stmt]
  (let [
    schema-info (load-schema-info sch-file)
    select-stmt (json/read-json select-stmt)
    attrs (let [attrs (:attributes select-stmt)] (if (empty? attrs) (map name (keys (schema-info :columnIndex))) attrs))
    where-clause (:whereClause select-stmt)]
    ; 存在しない属性をSELECT句に指定した場合
    (let [diff (set/difference (set attrs) (set (map name (keys (schema-info :columnIndex)))))]
      (if (not (empty? diff)) (throw (RuntimeException. (str "Attributes (" diff ") not exist in select statement.")))))
    ; 存在しない属性をWHERE句に指定した場合
    (let [diff (set/difference (set (map name (keys where-clause))) (set (map name (keys (schema-info :columnIndex)))))]
      (if (not (empty? diff)) (throw (RuntimeException. (str "Attributes (" diff ") not exist in where clause.")))))
    (json/write-str
      (if (empty? attrs)
        []
        (with-open [in (FileInputStream. xls-file)]
          (let [
            workbook (HSSFWorkbook. in)
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

(defn -selectSS [sch-file xls-file attrs]
  (selectSS sch-file xls-file attrs))

(defn getSSCellValues [xls-file sheet-idx addrs]
  (if (not (is-valid-cell-addr-lst addrs))
    (throw (RuntimeException. (str "Invalid cell address list. (" addrs ")")))
    (let [addrs (json/read-json addrs)]
      (json/write-str
        (with-open [in (FileInputStream. xls-file)]
          (let [
            workbook (HSSFWorkbook. in)
            sheet (.getSheetAt workbook sheet-idx)]
            (map (fn [addr] (get-cell-value sheet (first addr) (second addr))) addrs)))))))

(defn -getSSCellValues [xls-file sheet-idx addrs]
  (getSSCellValues xls-file sheet-idx addrs))

(defn setSSCellValues [xls-file sheet-idx kvs]
  (if (not (is-valid-cell-addr-val-lst kvs))
    (throw (RuntimeException. (str "Invalid cell address and value list. (" kvs ")")))
    (let [
      kvs (json/read-json kvs)
      in (FileInputStream. xls-file)]
      (try
        (let [
          workbook (HSSFWorkbook. in)
          sheet (.getSheetAt workbook sheet-idx)]
          (doseq [kv kvs] (set-cell-value sheet (nth kv 0) (nth kv 1) (nth kv 2)))
          (with-open [out (FileOutputStream. xls-file)]
            (.write workbook out)))
        (finally
          (.close in))))))

(defn -setSSCellValues [xls-file sheet-idx kvs]
  (setSSCellValues xls-file sheet-idx kvs))

(defn insertSS [sch-file xls-file key-value-map-set]
  (let [
    schema-info (load-schema-info sch-file)
    {
      col-idx-map :columnIndex
      stt-row-idx :startRowIndex
      end-row-idx :endRowIndex
      sheet-idx :sheetIndex
      required :required
    } schema-info
    kvms (json/read-json key-value-map-set)
    in (FileInputStream. xls-file)]
    (try
      (let [
        workbook (HSSFWorkbook. in)
        sheet (.getSheetAt workbook sheet-idx)]
        (letfn [
          ; 値を格納するセルアドレスを付与する
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
                  ; 必須属性が不足、存在しない属性を指定した場合は例外とする
                  (if
                    (not
                      (and
                        (empty? (set/difference req-set usr-set))
                        (empty? (set/difference usr-set col-set))))
                    (throw
                      (RuntimeException.
                        (str "Record (" kvm ") is not consistent with schema definition in the file (" sch-file ").")))))))
                kvms))
            avl-row-idxs (filter
              (complement (partial exist-required-value schema-info sheet))
              (range stt-row-idx (inc end-row-idx)))]
            ; 空き行が足りない場合は例外とする
            (if (> (count valid-kvms) (count avl-row-idxs))
              (throw
                (RuntimeException.
                  (str (str (count valid-kvms)) " rows insert failed (all row). Available row count is " (str (count avl-row-idxs)) "."))))                
            (doseq [kv (mapcat #(gen-addr-val-map %1 %2) valid-kvms avl-row-idxs)]
              (set-cell-value sheet (nth kv 0) (nth kv 1) (nth kv 2)))))
        (with-open [out (FileOutputStream. xls-file)]
          (.write workbook out)))
      (finally
        (.close in)))))

(defn -insertSS [sch-file xls-file key-value-map-set]
  (insertSS sch-file xls-file key-value-map-set))

(defn updateSS [sch-file xls-file update-stmts]
  (let [
    schema-info (load-schema-info sch-file)
    {
      col-idx-map :columnIndex
      stt-row-idx :startRowIndex
      end-row-idx :endRowIndex
      sheet-idx :sheetIndex
      required :required
    } schema-info
    update-stmts (json/read-json update-stmts)
    in (FileInputStream. xls-file)]
    (try
      (let [
        workbook (HSSFWorkbook. in)
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
                            (str "Record (" kvm ") is not consistent with schema definition in the file (" sch-file ").")))))))
                  update-stmts))]
            (doseq [kv (mapcat #(gen-addr-val-map-from-upd-stmt %1) valid-upd-stmts)]
              (set-cell-value sheet (nth kv 0) (nth kv 1) (nth kv 2)))))
        (with-open [out (FileOutputStream. xls-file)]
          (.write workbook out)))
      (finally
        (.close in)))))

(defn -updateSS [sch-file xls-file update-stmts]
  (updateSS sch-file xls-file update-stmts))