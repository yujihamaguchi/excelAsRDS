(ns excelAsRDS.DmlTest
  (:require
    [clojure.test :refer :all]
    [excelAsRDS.Dml :refer :all]
    [clojure.java.io :refer :all])
  (:import
    (org.apache.poi.ss.usermodel Workbook WorkbookFactory Name Row Cell)
    (org.apache.poi.ss.util AreaReference CellReference)
    (java.io FileInputStream FileOutputStream)))

(deftest work001
  (is
    (=
      (selectSS
        "./resources/SDM_TRM_IN_VIEW.json"
        "./resources/SDMインシデント.xls"
        "{ \"attributes\" : [
            \"status\"
            ] \"whereClause\" { \"r_persid\" : 6 }}"
      )
      "[]"
    )
  )
)
; (deftest work1
;   (is
;     (=
;       (selectSS
;         "./resources/SDM_TRM_IN_VIEW.json"
;         "./resources/tmp.xls"
;         "{ \"attributes\" : [
;             \"ref_num\"
;             , \"priority\"
;             , \"type\"
;             , \"customer\"
;             , \"active_flag\"
;             , \"zsupport_num\"
;             , \"zcmp_nm\"
;             , \"zprd_nm\"
;             , \"zbiz_kind\"
;             , \"zfun_kind\"
;             , \"zsubsystem\"
;             , \"zversion\"
;             , \"group_id\"
;             , \"category\"
;             , \"status\"
;             , \"log_agent\"
;             , \"requested_by\"
;             , \"zrequested_by_free\"
;             , \"zreq_email_address_free\"
;             , \"zreq_phone_number_free\"
;             , \"zreq_dept_nm_free\"
;             , \"zseverity\"
;             , \"zreq_res_date\"
;             , \"summary\"
;             , \"zoccurred_date\"
;             , \"description\"
;             , \"zbiz_kind_csv\"
;             , \"zfun_etc\"
;             , \"zprd_attr_csv\"
;             , \"zreceipt\"
;             , \"zrequest_type\"
;             , \"zin_src\"
;             , \"zserious_in_flg\"        
;         ] \"whereClause\" { \"load_flg\" : \"○\" }}")
;       "[]"
;     )
;   )
; )

; (deftest work2
;   (is
;     (=
;       (selectSS
;         "./resources/SDM_TRM_IN_VIEW.json"
;         "./resources/tmp.xls"
; "{ \"attributes\" : [
;     \"ref_num\"
;     , \"assignee\"
;     , \"urgency\"
;     , \"impact\"
;     , \"zdateline_date\"
;     , \"zinc_cnt_stat\"
;     , \"zenvironment\"
;     , \"znum_free\"
;     , \"zresult\"
;     , \"zcost\"
;     , \"zpost_review_flg\"
;     , \"zpost_review_fin_flg\"
;     , \"zapprover\"
;     , \"zapprover_free\"
;     , \"zapp_email_address_free\"
;     , \"zapp_phone_number_free\"
;     , \"zapp_dept_nm_free\"
;     , \"open_date\"
;     , \"last_mod_dt\"
;     , \"resolve_date\"
;     , \"close_date\"
;     , \"zdev_num\"
;     , \"zdev_kind\"
;     , \"zsrc_kind\"
;     , \"zprim_solve_flg\"
;     , \"zprim_solver_nm\"
;     , \"zworkaround_solve_flg\"
;     , \"last_mod_by\"
; ] \"whereClause\" { \"load_flg\" : \"○\" }}")
;       "[]"
;     )
;   )
; )


; (deftest tmp0
;   (let [
;     wb (WorkbookFactory/create (FileInputStream. "./resources/tmp.xls"))
;     sheet (.getSheetAt wb 1)
;     zreq_res_date (get-cell-value sheet 73 7 "TEXT(MONTH(P_ROWIDX_), \"00\")&\"/\"&TEXT(DAY(P_ROWIDX_), \"00\")&\"/\"&YEAR(P_ROWIDX_)&\" \"&TEXT(HOUR(P_ROWIDX_), \"00\")&\":\"&TEXT(MINUTE(P_ROWIDX_), \"00\")&\":\"&TEXT(SECOND(P_ROWIDX_), \"00\")")
;     ]
;     (is
;       (= zreq_res_date "D42CEC859D4F61498EA21CF4876C65B2")
;     )
;   )
; )

(deftest ut-get-cell-value-with-formula
  (testing "get-cell-value(正常系)"
    (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test14.xls")) 0)]
      (testing "文字列"
        (is
          (=
            (get-cell-value sheet 2 3 "A_ROWIDX_&B_ROWIDX_")
            "xx"
          )
        )
      )
      (testing "空文字"  
        (is
          (=
            (get-cell-value sheet 2 6 "A_ROWIDX_&B_ROWIDX_")
            ""
          )
        )
      )
      (testing "数値"
        (is
          (==
            (get-cell-value sheet 2 7 "A_ROWIDX_*B_ROWIDX_")
            300
          )
        )
      )
      (testing "日付"
        (is
          (=
            (get-cell-value sheet 2 8 "A_ROWIDX_+1")
            "2013/11/14"
          )
        )
      )
    )
  )
)

; (deftest tmp1
;   (let [
;     wb (WorkbookFactory/create (FileInputStream. "./resources/tmp.xls"))
;     sheet (.getSheetAt wb 1)
;     customer (do (set-cell-formula sheet 64 7 "IF(A8<>\"\",INDEX(acc!$C:$C,MATCH(A8,acc!$B:$B,0)),\"\")") (get-cell-value sheet 64 7))
;     active_flag (do (set-cell-formula sheet 65 7 "IF(B8,\"1\",\"0\")") (get-cell-value sheet 65 7))
;     zsupport_num (do (set-cell-formula sheet 66 7 "C8") (get-cell-value sheet 66 7))
;     zcmp_nm (do (set-cell-formula sheet 67 7 "D8") (get-cell-value sheet 67 7))
;     zprd_nm (do (set-cell-formula sheet 68 7 "E8") (get-cell-value sheet 68 7))
;     zbiz_kind (do (set-cell-formula sheet 69 7 "F8") (get-cell-value sheet 69 7))
;     zfun_kind (do (set-cell-formula sheet 70 7 "G8") (get-cell-value sheet 70 7))
;     zsubsystem (do (set-cell-formula sheet 71 7 "H8") (get-cell-value sheet 71 7))
;     zversion (do (set-cell-formula sheet 72 7 "I8") (get-cell-value sheet 72 7))
;     group_id (do (set-cell-formula sheet 73 7 "IF(J8<>\"\",INDEX(acc!$C:$C,MATCH(J8,acc!$B:$B,0)),\"\")") (get-cell-value sheet 73 7))
;     category (do (set-cell-formula sheet 74 7 "IF(M8<>\"\",INDEX(acc!$D:$D,MATCH(M8,acc!$B:$B,0)),\"\")") (get-cell-value sheet 74 7))
;     status (do (set-cell-formula sheet 75 7 "IF(N8<>\"\",INDEX(mst!$C:$C,MATCH(\"cr_stat|\"&N8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 75 7))
;     log_agent (do (set-cell-formula sheet 76 7 "IF(O8<>\"\",INDEX(acc!$C:$C,MATCH(O8,acc!$B:$B,0)),\"\")") (get-cell-value sheet 76 7))
;     requested_by (do (set-cell-formula sheet 77 7 "IF(P8<>\"\",INDEX(acc!$C:$C,MATCH(P8,acc!$B:$B,0)),\"\")") (get-cell-value sheet 77 7))
;     zrequested_by_free (do (set-cell-formula sheet 78 7 "IF(Q8<>\"\",Q8,\"\")") (get-cell-value sheet 78 7))
;     zreq_email_address_free (do (set-cell-formula sheet 79 7 "IF(R8<>\"\",R8,\"\")") (get-cell-value sheet 79 7))
;     zreq_phone_number_free (do (set-cell-formula sheet 80 7 "IF(S8<>\"\",S8,\"\")") (get-cell-value sheet 80 7))
;     zreq_dept_nm_free (do (set-cell-formula sheet 81 7 "IF(T8<>\"\",T8,\"\")") (get-cell-value sheet 81 7))
;     zseverity (do (set-cell-formula sheet 82 7 "IF(U8<>\"\",INDEX(mst!$C:$C,MATCH(\"zseverity|\"&U8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 82 7))
;     zreq_res_date "10/02/2012 00:00:00"
;     summary (do (set-cell-formula sheet 84 7 "W8") (get-cell-value sheet 84 7))
;     zoccurred_date "10/02/2012 00:00:00"
;     description (do (set-cell-formula sheet 86 7 "Y8") (get-cell-value sheet 86 7))
;     zbiz_kind_csv (do (set-cell-formula sheet 87 7 "IF(Z8<>\"\",INDEX(mst!$C:$C,MATCH(\"zbiz_kind_tree_list|\"&Z8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 87 7))
;     zfun_etc (do (set-cell-formula sheet 88 7 "IF(AA8<>\"\",AA8,\"\")") (get-cell-value sheet 88 7))
;     zprd_attr_csv (do (set-cell-formula sheet 89 7 "IF(AB8<>\"\",INDEX(mst!$C:$C,MATCH(\"zprd_attr_tree_list|\"&AB8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 89 7))
;     zreceipt (do (set-cell-formula sheet 90 7 "IF(AC8<>\"\",INDEX(mst!$C:$C,MATCH(\"zreceipt|\"&AC8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 90 7))
;     zrequest_type (do (set-cell-formula sheet 91 7 "IF(AD8<>\"\",INDEX(mst!$C:$C,MATCH(\"zrequest_type|\"&AD8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 91 7))
;     zin_src (do (set-cell-formula sheet 92 7 "IF(AE8<>\"\",INDEX(mst!$C:$C,MATCH(\"zin_src|\"&AE8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 92 7))
;     zserious_in_flg (do (set-cell-formula sheet 93 7 "IF(AF8<>\"\",IF(AF8,\"1\",\"0\"),\"\")") (get-cell-value sheet 93 7))
;     assignee (do (set-cell-formula sheet 94 7 "IF(AG8<>\"\",INDEX(acc!$C:$C,MATCH(AG8,acc!$B:$B,0)),\"\")") (get-cell-value sheet 94 7))
;     urgency (do (set-cell-formula sheet 95 7 "IF(AH8<>\"\",INDEX(mst!$C:$C,MATCH(\"urgncy|\"&AH8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 95 7))
;     impact (do (set-cell-formula sheet 96 7 "IF(AI8<>\"\",INDEX(mst!$C:$C,MATCH(\"impact|\"&AI8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 96 7))
;     zdateline_date "10/02/2012 00:00:00"
;     zinc_cnt_stat (do (set-cell-formula sheet 98 7 "IF(AK8<>\"\",INDEX(mst!$C:$C,MATCH(\"zinc_cnt_stat|\"&AK8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 98 7))
;     zenvironment (do (set-cell-formula sheet 99 7 "IF(AL8<>\"\",INDEX(mst!$C:$C,MATCH(\"zenvironment|\"&AL8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 99 7))
;     znum_free (do (set-cell-formula sheet 100 7 "IF(AM8<>\"\",AM8,\"\")") (get-cell-value sheet 100 7))
;     zresult (do (set-cell-formula sheet 101 7 "IF(AN8<>\"\",AN8,\"\")") (get-cell-value sheet 101 7))
;     zcost (do (set-cell-formula sheet 102 7 "IF(AO8<>\"\",AO8,\"\")") (get-cell-value sheet 102 7))
;     zpost_review_flg (do (set-cell-formula sheet 103 7 "IF(AP8<>\"\",IF(AP8,\"1\",\"0\"),\"\")") (get-cell-value sheet 103 7))
;     zpost_review_fin_flg (do (set-cell-formula sheet 104 7 "IF(AQ8<>\"\",IF(AQ8,\"1\",\"0\"),\"\")") (get-cell-value sheet 104 7))
;     zapprover (do (set-cell-formula sheet 105 7 "IF(AR8<>\"\",INDEX(acc!$C:$C,MATCH(AR8,acc!$B:$B,0)),\"\")") (get-cell-value sheet 105 7))
;     zapprover_free (do (set-cell-formula sheet 106 7 "IF(AS8<>\"\",AS8,\"\")") (get-cell-value sheet 106 7))
;     zapp_email_address_free (do (set-cell-formula sheet 107 7 "IF(AT8<>\"\",AT8,\"\")") (get-cell-value sheet 107 7))
;     zapp_phone_number_free (do (set-cell-formula sheet 108 7 "IF(AU8<>\"\",AU8,\"\")") (get-cell-value sheet 108 7))
;     zapp_dept_nm_free (do (set-cell-formula sheet 109 7 "IF(AV8<>\"\",AV8,\"\")") (get-cell-value sheet 109 7))
;     open_date "10/02/2012 00:00:00"
;     last_mod_dt "10/02/2012 00:00:00"
;     resolve_date "10/02/2012 00:00:00"
;     close_date "10/02/2012 00:00:00"
;     zdev_num (do (set-cell-formula sheet 114 7 "IF(BA8<>\"\",BA8,\"\")") (get-cell-value sheet 114 7))
;     zdev_kind (do (set-cell-formula sheet 115 7 "IF(BB8<>\"\",INDEX(mst!$C:$C,MATCH(\"zdev_kind|\"&BB8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 115 7))
;     zsrc_kind (do (set-cell-formula sheet 116 7 "IF(BC8<>\"\",INDEX(mst!$C:$C,MATCH(\"zsrc_kind|\"&BC8,mst!$D:$D,0)),\"\")") (get-cell-value sheet 116 7))
;     zprim_solve_flg (do (set-cell-formula sheet 120 7 "IF(BG8<>\"\",IF(BG8,\"1\",\"0\"),\"\")") (get-cell-value sheet 120 7))
;     zprim_solver_nm (do (set-cell-formula sheet 121 7 "IF(BH8<>\"\",INDEX(acc!$C:$C,MATCH(BH8,acc!$B:$B,0)),\"\")") (get-cell-value sheet 121 7))
;     zworkaround_solve_flg (do (set-cell-formula sheet 122 7 "IF(BI8<>\"\",IF(BI8,\"1\",\"0\"),\"\")") (get-cell-value sheet 122 7))
;     last_mod_by (do (set-cell-formula sheet 123 7 "IF(BJ8<>\"\",INDEX(acc!$C:$C,MATCH(BJ8,acc!$B:$B,0)),\"\")") (get-cell-value sheet 123 7))
;     ]
;     (is (= customer "C271D1C1553F3341A9B61A6556F4880F"))
;     (is (= active_flag "1"))
;     (is (= zsupport_num "114301"))
;     (is (= zcmp_nm "日本トリプルウィン"))
;     (is (= zprd_nm "POSITIVE"))
;     (is (= zbiz_kind "人事"))
;     (is (= zfun_kind "異動管理"))
;     (is (= zsubsystem "01.POSITIVE"))
;     (is (= zversion "3.1"))
;     (is (= group_id "C271D1C1553F3341A9B61A6556F4880F"))
;     (is (= category "pcat:400072"))
;     (is (= status "OP"))
;     (is (= log_agent "BE8B3DCA8C50E443970E2607A0648569"))
;     (is (= requested_by "743FD4B4021E0C428D75CCF9B92362B0"))
;     (is (= zrequested_by_free "依頼者（手入力）"))
;     (is (= zreq_email_address_free "依頼者連絡先E-mail（手入力）"))
;     (is (= zreq_phone_number_free "依頼者連絡先TEL（手入力）"))
;     (is (= zreq_dept_nm_free "依頼者部署名（手入力）"))
;     (is (= zseverity "400002"))
;     (is (= zreq_res_date "10/02/2012 00:00:00"))
;     (is (= summary "件名"))
;     (is (= zoccurred_date "10/02/2012 00:00:00"))
;     (is (= description "内容"))
;     (is (= zbiz_kind_csv "400073"))
;     (is (= zfun_etc "その他機能"))
;     (is (= zprd_attr_csv "400073"))
;     (is (= zreceipt "400005"))
;     (is (= zrequest_type "400001"))
;     (is (= zin_src "400002"))
;     (is (= zserious_in_flg "1"))
;     (is (= assignee "78F7273D7D7A1245BDA887A046C22352"))
;     (is (= urgency "1"))
;     (is (= impact "3"))
;     (is (= zdateline_date "10/02/2012 00:00:00"))
;     (is (= zinc_cnt_stat "400001"))
;     (is (= zenvironment "400002"))
;     (is (= znum_free "個別管理番号"))
;     (is (= zresult "作業結果・最終回答"))
;     (is (= zcost 123.0))
;     (is (= zpost_review_flg "1"))
;     (is (= zpost_review_fin_flg "1"))
;     (is (= zapprover "9A976C3A5B56204C91FF166BD1DF9EC1"))
;     (is (= zapprover_free "クローズ承認者（手入力）"))
;     (is (= zapp_email_address_free "クローズ承認者連絡先E-mail（手入力）"))
;     (is (= zapp_phone_number_free "クローズ承認者連絡先TEL（手入力）"))
;     (is (= zapp_dept_nm_free "クローズ承認者部署名（手入力）"))
;     (is (= open_date "10/02/2012 00:00:00"))
;     (is (= last_mod_dt "10/02/2012 00:00:00"))
;     (is (= resolve_date "10/02/2012 00:00:00"))
;     (is (= close_date "10/02/2012 00:00:00"))
;     (is (= zdev_num "開発管理番号"))
;     (is (= zdev_kind "400001"))
;     (is (= zsrc_kind "400003"))
;     (is (= zprim_solve_flg "1"))
;     (is (= zprim_solver_nm "A1CC699820B6EB4C90FB7A13375D7CCA"))
;     (is (= zworkaround_solve_flg "1"))
;     (is (= last_mod_by "D42CEC859D4F61498EA21CF4876C65B2"))
;   )
; )

(deftest ut-get-cell-value
  (testing "get-cell-value(正常系)"
    (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test01.xls")) 0)]
      (testing "文字列"
        (is
          (=
            (get-cell-value sheet 0 3)
            "x"
          )
        )
      )
      (testing "空文字"  
        (is
          (=
            (get-cell-value sheet 100 0)
            ""
          )
        )
      )
      (testing "数値"
        (is
          (==
            (get-cell-value sheet 0 1)
            1
          )
        )
      )
      (testing "日付"
        (is
          (=
            (get-cell-value sheet 1 1)
            "2013/07/29"
          )
        )
      )
    )
  )
)

; (deftest ut-get-cell-value-xlsx
;   (testing "get-cell-value(正常系)"
;     (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test01.xlsx")) 0)]
;       (testing "文字列"
;         (is
;           (=
;             (get-cell-value sheet 0 3)
;             "x"
;           )
;         )
;       )
;       (testing "空文字"  
;         (is
;           (=
;             (get-cell-value sheet 100 0)
;             ""
;           )
;         )
;       )
;       (testing "数値"
;         (is
;           (==
;             (get-cell-value sheet 0 1)
;             1
;           )
;         )
;       )
;       (testing "日付"
;         (is
;           (=
;             (get-cell-value sheet 1 1)
;             "2013/07/29"
;           )
;         )
;       )
;     )
;   )
; )

(deftest ut-get-cell-value-ab
  (testing "get-cell-value(異常系)"
    (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test01.xls")) 0)]
      (testing "行が範囲外1"
        (is
          (=
            (try (get-cell-value sheet 0 -1) (catch Exception e (.getMessage e)))
            ""
          )
        )
      )
      (testing "行が範囲外2"
        (is
          (=
            (try (get-cell-value sheet 0 65536) (catch Exception e (.getMessage e)))
            ""
          )
        )
      )
      (testing "列が範囲外1"
        (is
          (=
            (get-cell-value sheet -1 3)
            ""
          )
        )
      )
      (testing "列が範囲外2"
        (is
          (=
            (get-cell-value sheet 256 3)
            ""
          )
        )
      )
    )
  )
)

; (deftest ut-get-cell-value-ab-xlsx
;   (testing "get-cell-value(異常系)"
;     (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test01.xlsx")) 0)]
;       (testing "行が範囲外1"
;         (is
;           (=
;             (try (get-cell-value sheet 0 -1) (catch Exception e (.getMessage e)))
;             ""
;           )
;         )
;       )
;       (testing "行が範囲外2"
;         (is
;           (=
;             (try (get-cell-value sheet 0 65536) (catch Exception e (.getMessage e)))
;             ""
;           )
;         )
;       )
;       (testing "列が範囲外1"
;         (is
;           (=
;             (try (get-cell-value sheet -1 3) (catch Exception e (.getMessage e)))
;             "Cell index must be >= 0"
;           )
;         )
;       )
;       (testing "列が範囲外2"
;         (is
;           (=
;             (get-cell-value sheet 16384 3)
;             ""
;           )
;         )
;       )
;     )
;   )
; )

(deftest ut-meet-where-clause-cond
  (testing "meet-where-clause-cond(正常系)"
    (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test01.xls")) 0)]
      (testing "整数"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 10 {:attr1 1})
            true
          )
        )
      )
      (testing "小数"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 10 {:attr2 1.1})
            true
          )
        )
      )
      (testing "文字列"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 11 {:attr1 "abc"})
            true
          )
        )
      )
      (testing "文字列(マルチバイト)"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 11 {:attr2 "あいう"})
            true
          )
        )
      )
      (testing "日付"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 12 {:attr1 "2013/08/10"})
            true
          )
        )
      )
      (testing "複合条件1"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 13 {:attr1 1 :attr2 "abc"})
            true
          )
        )
      )
      (testing "複合条件2"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 13 {:attr1 2 :attr2 "abc"})
            false
          )
        )
      )
      (testing "複合条件3"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 13 {:attr1 1 :attr2 "xyz"})
            false
          )
        )
      )
      (testing "複合条件4"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 14 {:attr1 "x" :attr2 "xxx" :attr3 "z"})
            false
          )
        )
      )
      (testing "Where句の指定が無い"
        (is
          (=
            (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 10 nil)
            true
          )
        )
      )
    )
  )
)

; (deftest ut-meet-where-clause-cond-xlsx
;   (testing "meet-where-clause-cond(正常系)"
;     (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test01.xlsx")) 0)]
;       (testing "整数"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 10 {:attr1 1})
;             true
;           )
;         )
;       )
;       (testing "小数"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 10 {:attr2 1.1})
;             true
;           )
;         )
;       )
;       (testing "文字列"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 11 {:attr1 "abc"})
;             true
;           )
;         )
;       )
;       (testing "文字列(マルチバイト)"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 11 {:attr2 "あいう"})
;             true
;           )
;         )
;       )
;       (testing "日付"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 12 {:attr1 "2013/08/10"})
;             true
;           )
;         )
;       )
;       (testing "複合条件1"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 13 {:attr1 1 :attr2 "abc"})
;             true
;           )
;         )
;       )
;       (testing "複合条件2"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 13 {:attr1 2 :attr2 "abc"})
;             false
;           )
;         )
;       )
;       (testing "複合条件3"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 13 {:attr1 1 :attr2 "xyz"})
;             false
;           )
;         )
;       )
;       (testing "複合条件4"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 14 {:attr1 "x" :attr2 "xxx" :attr3 "z"})
;             false
;           )
;         )
;       )
;       (testing "Where句の指定が無い"
;         (is
;           (=
;             (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 10 nil)
;             true
;           )
;         )
;       )
;     )
;   )
; )

(deftest ut-meet-where-clause-cond-ab
  (testing "meet-where-clause-cond(異常系)"
    (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test01.xls")) 0)]
      (testing "列アドレスの指定が無い"
        (is
          (=
            (try (meet-where-clause-cond {:columnIndex {}} sheet 10 {:attr1 1}) (catch Exception e (.getName (class e))))
            "java.lang.IllegalArgumentException"
          )
        )
      )
      (testing "Where句に存在しない属性を指定"
        (is
          (=
            (try (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 10 {:attr9 1}) (catch Exception e (.getName (class e))))
            "java.lang.IllegalArgumentException"
          )
        )
      )
    )
  )
)

; (deftest ut-meet-where-clause-cond-ab-xlsx
;   (testing "meet-where-clause-cond(異常系)"
;     (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test01.xlsx")) 0)]
;       (testing "列アドレスの指定が無い"
;         (is
;           (=
;             (try (meet-where-clause-cond {:columnIndex {}} sheet 10 {:attr1 1}) (catch Exception e (.getName (class e))))
;             "java.lang.IllegalArgumentException"
;           )
;         )
;       )
;       (testing "Where句に存在しない属性を指定"
;         (is
;           (=
;             (try (meet-where-clause-cond {:columnIndex {:attr1 0 :attr2 1 :attr3 2}} sheet 10 {:attr9 1}) (catch Exception e (.getName (class e))))
;             "java.lang.IllegalArgumentException"
;           )
;         )
;       )
;     )
;   )
; )

(deftest ut-exist-required-value
  (testing "exist-required-value(正常系)"
    (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test02.xls")) 0)]
      (testing "単独キー＆値が存在する"
        (is
          (=
            (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1"]} sheet 1)
            true
          )
        )
      )
      (testing "単独キー＆値が存在しない"
        (is
          (=
            (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1"]} sheet 3)
            false
          )
        )
      )
      (testing "複合キー（2d）＆値が一部存在しない"
        (is
          (=
            (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1" "attr2"]} sheet 1)
            false
          )
        )
      )
      (testing "複合キー（2d）＆値が存在する"
        (is
          (=
            (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1" "attr2"]} sheet 2)
            true
          )
        )
      )
      (testing "複合キー（2d）＆値が全て存在しない"
        (is
          (=
            (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1" "attr2"]} sheet 3)
            false
          )
        )
      )
      (testing "複合キー（3d）＆値が一部存在しない"
        (is
          (=
            (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1" "attr2" "attr3"]} sheet 2)
            false
          )
        )
      )
      (testing "必須属性無し1"
        (is
          (=
            (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required []} sheet 1)
            true
          )
        )
      )
      (testing "必須属性無し2"
        (is
          (=
            (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required []} sheet 2)
            true
          )
        )
      )
      (testing "必須属性無し3"
        (is
          (=
            (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required []} sheet 3)
            true
          )
        )
      )
    )
  )
)

; (deftest ut-exist-required-value-xlsx
;   (testing "exist-required-value(正常系)"
;     (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test02.xlsx")) 0)]
;       (testing "単独キー＆値が存在する"
;         (is
;           (=
;             (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1"]} sheet 1)
;             true
;           )
;         )
;       )
;       (testing "単独キー＆値が存在しない"
;         (is
;           (=
;             (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1"]} sheet 3)
;             false
;           )
;         )
;       )
;       (testing "複合キー（2d）＆値が一部存在しない"
;         (is
;           (=
;             (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1" "attr2"]} sheet 1)
;             false
;           )
;         )
;       )
;       (testing "複合キー（2d）＆値が存在する"
;         (is
;           (=
;             (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1" "attr2"]} sheet 2)
;             true
;           )
;         )
;       )
;       (testing "複合キー（2d）＆値が全て存在しない"
;         (is
;           (=
;             (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1" "attr2"]} sheet 3)
;             false
;           )
;         )
;       )
;       (testing "複合キー（3d）＆値が一部存在しない"
;         (is
;           (=
;             (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1" "attr2" "attr3"]} sheet 2)
;             false
;           )
;         )
;       )
;       (testing "必須属性無し1"
;         (is
;           (=
;             (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required []} sheet 1)
;             true
;           )
;         )
;       )
;       (testing "必須属性無し2"
;         (is
;           (=
;             (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required []} sheet 2)
;             true
;           )
;         )
;       )
;       (testing "必須属性無し3"
;         (is
;           (=
;             (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required []} sheet 3)
;             true
;           )
;         )
;       )
;     )
;   )
; )

(deftest ut-exist-required-value-ab
  (testing "exist-required-value(異常系)"
    (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test02.xls")) 0)]
    (testing "行が範囲外1"
      (is
        (=
          (try (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1"]} sheet -1) (catch Exception e (.getMessage e)))
          false
        )
      )
    )
    (testing "行が範囲外2"
      (is
        (=
          (try (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1"]} sheet 65536) (catch Exception e (.getMessage e)))
          false
          )
        )
      )
    )
  )
)

; (deftest ut-exist-required-value-ab-xlsx
;   (testing "exist-required-value(異常系)"
;     (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test02.xlsx")) 0)]
;     (testing "行が範囲外1"
;       (is
;         (=
;           (try (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1"]} sheet -1) (catch Exception e (.getMessage e)))
;           false
;         )
;       )
;     )
;     (testing "行が範囲外2"
;       (is
;         (=
;           (try (exist-required-value {:columnIndex {:attr1 0 :attr2 1 :attr3 2} :required ["attr1"]} sheet 65536) (catch Exception e (.getMessage e)))
;           false
;           )
;         )
;       )
;     )
;   )
; )

(deftest ut-load-schema-info
  (testing "load-schema-info(正常系)"
    (testing "通常処理"
      (is
        (=
          (load-schema-info "./resources/test01.json")
          {:sheetIndex 0, :columnIndex {:id 0, :pwd 1}, :startRowIndex 3, :endRowIndex 5, :required ["id"]}
        ) 
      )
    )
  )
)

(deftest ut-load-schema-info-ab
  (testing "load-schema-info(異常系)"
    (testing "システム必須属性が不足している"
      (is
        (let [schema-file-name-name "./resources/test03.json"]
          (=
            (try (load-schema-info schema-file-name-name) (catch Exception e (.getMessage e)))
            (str "Required attributes (#{:endRowIndex :startRowIndex}) not exist in file (" schema-file-name-name ").")
          )
        )
      )
    )
    (testing "列アドレスが未定義"
      (is
        (let [schema-file-name-name "./resources/test06.json"]
          (=
            (try (load-schema-info schema-file-name-name) (catch Exception e (.getMessage e)))
            (str "Column index definition (key 'columnIndex') not exist in file (" schema-file-name-name ").")
          )
        )
      )
    )
    (testing "ファイルが存在しない"
      (is
        (=
          (try (load-schema-info "./resources/not-exists.json") (catch Exception e (.getName (class e))))
          "java.io.FileNotFoundException"
        )
      )
    )
    (testing "ファイルを解析できない"
      (is
        (let [
          errMessage (try (load-schema-info "./resources/test05.json") (catch Exception e (.getMessage e)))]
          (re-find #"^JSON error" errMessage)
        )
      )
    )
  )
)

(deftest ut-selectSS
  (testing "selectSS(正常系)"
    (testing "2属性指定、必須属性完備"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] }")
          "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"x\",\"pwd\":\"p1\"},{\"id\":\"z\",\"pwd\":\"p3\"}]"
        )
      )
    )
    (testing "Where句指定1"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"x\" }}")
          "[{\"id\":\"x\",\"pwd\":\"p1\"}]"
        )
      )
    )
    (testing "Where句指定2"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"y\"}}")
          "[{\"id\":\"y\",\"pwd\":\"p2\"}]"
        )
      )
    )
    (testing "Where句指定3"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"pwd\" : \"p3\"}}")
          "[{\"id\":\"z\",\"pwd\":\"p3\"}]"
        )
      )
    )
    (testing "Where句指定4"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"x\", \"pwd\" : \"p1\"}}")
          "[{\"id\":\"x\",\"pwd\":\"p1\"}]"
        )
      )
    )
    (testing "Where句指定5"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"xx\"}}")
          "[]"
        )
      )
    )
    (testing "Where句指定6"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"x\", \"pwd\" : \"p2\"}}")
          "[]"
        )
      )
    )
    (testing "2属性指定、1件必須属性無し"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test03.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] }")
          "[{\"id\":\"x\",\"pwd\":\"p1\"},{\"id\":\"z\",\"pwd\":\"p3\"}]"
        )
      )
    )
    (testing "3属性、1件必須属性無し"
      (is
        (=
          (selectSS
            "./resources/test02.json"
            "./resources/test04.xls"
            "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
          "[{\"host\":\"z\",\"id\":\"p3\",\"pwd\":\"pppz\"},{\"host\":\"x\",\"id\":\"p1\",\"pwd\":\"pppx\"}]"
        )
      )
    )
    (testing "属性指定がない"
      (is
        (=
          (selectSS
            "./resources/test02.json"
            "./resources/test04.xls"
            "{ \"attributes\" : [] }")
          "[{\"host\":\"z\",\"id\":\"p3\",\"pwd\":\"pppz\"},{\"host\":\"x\",\"id\":\"p1\",\"pwd\":\"pppx\"}]"
        )
      )
    )
    (testing "属性指定がない、Where句指定はある"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "{ \"whereClause\" { \"id\" : \"y\" }}")
          "[{\"id\":\"y\",\"pwd\":\"p2\"}]"
        )
      )
    )
    (testing "属性指定がないWhere句もない"
      (is
        (=
          (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "{}")
          "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"x\",\"pwd\":\"p1\"},{\"id\":\"z\",\"pwd\":\"p3\"}]"
        )
      )
    )
  )
)

; (deftest ut-selectSS-xlsx
;   (testing "selectSS(正常系)"
;     (testing "2属性指定、必須属性完備"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"x\",\"pwd\":\"p1\"},{\"id\":\"z\",\"pwd\":\"p3\"}]"
;         )
;       )
;     )
;     (testing "Where句指定1"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"x\" }}")
;           "[{\"id\":\"x\",\"pwd\":\"p1\"}]"
;         )
;       )
;     )
;     (testing "Where句指定2"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"y\"}}")
;           "[{\"id\":\"y\",\"pwd\":\"p2\"}]"
;         )
;       )
;     )
;     (testing "Where句指定3"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"pwd\" : \"p3\"}}")
;           "[{\"id\":\"z\",\"pwd\":\"p3\"}]"
;         )
;       )
;     )
;     (testing "Where句指定4"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"x\", \"pwd\" : \"p1\"}}")
;           "[{\"id\":\"x\",\"pwd\":\"p1\"}]"
;         )
;       )
;     )
;     (testing "Where句指定5"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"xx\"}}")
;           "[]"
;         )
;       )
;     )
;     (testing "Where句指定6"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"x\", \"pwd\" : \"p2\"}}")
;           "[]"
;         )
;       )
;     )
;     (testing "2属性指定、1件必須属性無し"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test03.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           "[{\"id\":\"x\",\"pwd\":\"p1\"},{\"id\":\"z\",\"pwd\":\"p3\"}]"
;         )
;       )
;     )
;     (testing "3属性、1件必須属性無し"
;       (is
;         (=
;           (selectSS
;             "./resources/test02.json"
;             "./resources/test04.xlsx"
;             "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
;           "[{\"host\":\"z\",\"id\":\"p3\",\"pwd\":\"pppz\"},{\"host\":\"x\",\"id\":\"p1\",\"pwd\":\"pppx\"}]"
;         )
;       )
;     )
;     (testing "属性指定がない"
;       (is
;         (=
;           (selectSS
;             "./resources/test02.json"
;             "./resources/test04.xlsx"
;             "{ \"attributes\" : [] }")
;           "[{\"host\":\"z\",\"id\":\"p3\",\"pwd\":\"pppz\"},{\"host\":\"x\",\"id\":\"p1\",\"pwd\":\"pppx\"}]"
;         )
;       )
;     )
;     (testing "属性指定がない、Where句指定はある"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "{ \"whereClause\" { \"id\" : \"y\" }}")
;           "[{\"id\":\"y\",\"pwd\":\"p2\"}]"
;         )
;       )
;     )
;     (testing "属性指定がないWhere句もない"
;       (is
;         (=
;           (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "{}")
;           "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"x\",\"pwd\":\"p1\"},{\"id\":\"z\",\"pwd\":\"p3\"}]"
;         )
;       )
;     )
;   )
; )

(deftest ut-selectSS-ab
  (testing "selectSS(異常系)"
    (testing "スキーマ定義ファイルにて、システム必須属性が不足している"
      (is
        (let [schema-file-name-name "./resources/test03.json"]
          (=
            (try (selectSS
              schema-file-name-name
              "./resources/test01.xls"
              "{ \"attributes\" : [\"id\", \"pwd\"] }")
            (catch Exception e (.getMessage e)))
            (str "Required attributes (#{:endRowIndex :startRowIndex}) not exist in file (" schema-file-name-name ").")
          )
        )
      )
    )
    (testing "スキーマ定義ファイルにて、列アドレスが未定義"
      (is
        (let [schema-file-name-name "./resources/test06.json"]
          (=
            (try (selectSS
              schema-file-name-name
              "./resources/test01.xls"
              "{ \"attributes\" : [\"id\", \"pwd\"] }")
            (catch Exception e (.getMessage e)))
            (str "Column index definition (key 'columnIndex') not exist in file (" schema-file-name-name ").")
          )
        )
      )
    )
    (testing "スキーマ定義ファイルが存在しない"
      (is
        (=
          (try (selectSS
            "./resources/not-exists.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] }")
          (catch Exception e (.getName (class e))))
          "java.io.FileNotFoundException"
        )
      )
    )
    (testing "スキーマ定義ファイルが解析出来ない"
      (is
        (let [
          errMessage (try (selectSS
            "./resources/test05.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] }")
          (catch Exception e (.getMessage e)))]
          (re-find #"^JSON error" errMessage))
      )
    )
    (testing "エクセルファイルが存在しない"
      (is
        (=
          (try (selectSS
            "./resources/test01.json"
            "X:/not-exists.file"
            "{ \"attributes\" : [\"id\", \"pwd\"] }")
          (catch Exception e (.getName (class e))))
          "java.io.FileNotFoundException"
        )
      )
    )
    (testing "エクセルファイルを解析できない"
      (is
        (=
          (try (selectSS
            "./resources/test01.json"
            "./resources/not-excel.txt"
            "{ \"attributes\" : [\"id\", \"pwd\"] }")
          (catch Exception e (.getName (class e))))
          "java.lang.IllegalArgumentException"
        )
      )
    )
    (testing "エクセルファイルの指定のシートが存在しない"
      (is
        (=
          (try (selectSS
            "./resources/test07.json"
            "./resources/test01.xls"
            "{ \"attributes\" : [\"id\", \"pwd\"] }")
          (catch Exception e (vector (.getName (class e)) (.getMessage e))))
          ["java.lang.IllegalArgumentException" "Sheet index (1) is out of range (0..0)"]
        )
      )
    )
    (testing "存在しない属性を指定"
      (is
        (=
          (try
            (selectSS
              "./resources/test02.json"
              "./resources/test04.xls"
              "{ \"attributes\" : [\"not-exist-attr\", \"host\", \"id\", \"pwd\"] }")
            (catch RuntimeException e (.getMessage e)))
          "Attributes (#{\"not-exist-attr\"}) not exist in select statement."
        )
      )
    )
    (testing "属性取得、条件指定文字列を解析できない1"
      (is
        (let [
          errMessage (try (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "")
          (catch Exception e (.getMessage e)))]
          (re-find #"^JSON error" errMessage))
      )
    )
    (testing "属性取得、条件指定文字列を解析できない2"
      (is
        (let [
          errMessage (try (selectSS
            "./resources/test01.json"
            "./resources/test01.xls"
            "xxx")
          (catch Exception e (.getMessage e)))]
          (re-find #"^JSON error" errMessage))
      )
    )
    (testing "存在しない属性をWhere句に指定"
      (is
        (=
          (try
            (selectSS
              "./resources/test02.json"
              "./resources/test04.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\"], \"whereClause\" : {\"not-exist-attr\" : 1 }}")
            (catch RuntimeException e (.getMessage e)))
          "Attributes (#{\"not-exist-attr\"}) not exist in where clause."
        )
      )
    )
  )
)

; (deftest ut-selectSS-ab-xlsx
;   (testing "selectSS(異常系)"
;     (testing "スキーマ定義ファイルにて、システム必須属性が不足している"
;       (is
;         (let [schema-file-name-name "./resources/test03.json"]
;           (=
;             (try (selectSS
;               schema-file-name-name
;               "./resources/test01.xlsx"
;               "{ \"attributes\" : [\"id\", \"pwd\"] }")
;             (catch Exception e (.getMessage e)))
;             (str "Required attributes (#{:endRowIndex :startRowIndex}) not exist in file (" schema-file-name-name ").")
;           )
;         )
;       )
;     )
;     (testing "スキーマ定義ファイルにて、列アドレスが未定義"
;       (is
;         (let [schema-file-name-name "./resources/test06.json"]
;           (=
;             (try (selectSS
;               schema-file-name-name
;               "./resources/test01.xls"
;               "{ \"attributes\" : [\"id\", \"pwd\"] }")
;             (catch Exception e (.getMessage e)))
;             (str "Column index definition (key 'columnIndex') not exist in file (" schema-file-name-name ").")
;           )
;         )
;       )
;     )
;     (testing "スキーマ定義ファイルが存在しない"
;       (is
;         (=
;           (try (selectSS
;             "./resources/not-exists.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           (catch Exception e (.getName (class e))))
;           "java.io.FileNotFoundException"
;         )
;       )
;     )
;     (testing "スキーマ定義ファイルが解析出来ない"
;       (is
;         (let [
;           errMessage (try (selectSS
;             "./resources/test05.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           (catch Exception e (.getMessage e)))]
;           (re-find #"^JSON error" errMessage))
;       )
;     )
;     (testing "エクセルファイルが存在しない"
;       (is
;         (=
;           (try (selectSS
;             "./resources/test01.json"
;             "X:/not-exists.file"
;             "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           (catch Exception e (.getName (class e))))
;           "java.io.FileNotFoundException"
;         )
;       )
;     )
;     (testing "エクセルファイルを解析できない"
;       (is
;         (=
;           (try (selectSS
;             "./resources/test01.json"
;             "./resources/not-excel.txt"
;             "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           (catch Exception e (.getName (class e))))
;           "java.lang.IllegalArgumentException"
;         )
;       )
;     )
;     (testing "エクセルファイルの指定のシートが存在しない"
;       (is
;         (=
;           (try (selectSS
;             "./resources/test07.json"
;             "./resources/test01.xlsx"
;             "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           (catch Exception e (vector (.getName (class e)) (.getMessage e))))
;           ["java.lang.IllegalArgumentException" "Sheet index (1) is out of range (0..0)"]
;         )
;       )
;     )
;     (testing "存在しない属性を指定"
;       (is
;         (=
;           (try
;             (selectSS
;               "./resources/test02.json"
;               "./resources/test04.xlsx"
;               "{ \"attributes\" : [\"not-exist-attr\", \"host\", \"id\", \"pwd\"] }")
;             (catch RuntimeException e (.getMessage e)))
;           "Attributes (#{\"not-exist-attr\"}) not exist in select statement."
;         )
;       )
;     )
;     (testing "属性取得、条件指定文字列を解析できない1"
;       (is
;         (let [
;           errMessage (try (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "")
;           (catch Exception e (.getMessage e)))]
;           (re-find #"^JSON error" errMessage))
;       )
;     )
;     (testing "属性取得、条件指定文字列を解析できない2"
;       (is
;         (let [
;           errMessage (try (selectSS
;             "./resources/test01.json"
;             "./resources/test01.xlsx"
;             "xxx")
;           (catch Exception e (.getMessage e)))]
;           (re-find #"^JSON error" errMessage))
;       )
;     )
;     (testing "存在しない属性をWhere句に指定"
;       (is
;         (=
;           (try
;             (selectSS
;               "./resources/test02.json"
;               "./resources/test04.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\"], \"whereClause\" : {\"not-exist-attr\" : 1 }}")
;             (catch RuntimeException e (.getMessage e)))
;           "Attributes (#{\"not-exist-attr\"}) not exist in where clause."
;         )
;       )
;     )
;   )
; )

(deftest ut-is-valid-cell-addr-coll
  (testing "is-valid-cell-addr-coll(正常系)"
    (is (= (is-valid-cell-addr-coll "[[1,2]]") true ) )
    (is (= (is-valid-cell-addr-coll "[[1,2],[3,4]]") true ) )
    (is (= (is-valid-cell-addr-coll "") false ) )
    (is (= (is-valid-cell-addr-coll "1") false ) )
    (is (= (is-valid-cell-addr-coll "[]") false ) )
    (is (= (is-valid-cell-addr-coll "[1]") false ) )
    (is (= (is-valid-cell-addr-coll "[[1,2],1]") false ) )
    (is (= (is-valid-cell-addr-coll "[1,[1,2]]") false ) )
    (is (= (is-valid-cell-addr-coll "[[1,2],[3,4,5]]") false ) )
    (is (= (is-valid-cell-addr-coll "[[1,2,3],[4,5]]") false ) )
    (is (= (is-valid-cell-addr-coll "[[1],[4,5]]") false ) )
    (is (= (is-valid-cell-addr-coll "[[1,2],[]]") false ) )
    (is (= (is-valid-cell-addr-coll "[1,2]") false ) )
    (is (= (is-valid-cell-addr-coll "[[]]") false ) )
    (is (= (is-valid-cell-addr-coll "[[1]]") false ) )
    (is (= (is-valid-cell-addr-coll "{}") false ) )
    (is (= (is-valid-cell-addr-coll "[[1,2],[3,\"x\"]]") false ) )
    (is (= (is-valid-cell-addr-coll "[[1,2],[\"x\",3]]") false ) )
    (is (= (is-valid-cell-addr-coll "[[1,\"x\"],[3,4]]") false ) )
    (is (= (is-valid-cell-addr-coll "[[\"x\",2],[3,4]]") false ) )
  )
)

(deftest ut-is-valid-cell-addr-val-coll
  (testing "is-valid-cell-addr-val-coll(正常系)"
    (is (= (is-valid-cell-addr-val-coll "[[1,2,3]]") true ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2,3],[3,4,\"x\"]]") true ) )
    (is (= (is-valid-cell-addr-val-coll "") false ) )
    (is (= (is-valid-cell-addr-val-coll "1") false ) )
    (is (= (is-valid-cell-addr-val-coll "[]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[1]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2,3],1]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2],[1,2,3]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2,3],[3,4,5,6]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2,3,4],[4,5,6]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2],[3,4,5]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2,3],[]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[1,2,3]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "{}") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2,3],[3,\"x\",4]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,2,3],[\"x\",3,4]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[1,\"x\",3],[3,4,5]]") false ) )
    (is (= (is-valid-cell-addr-val-coll "[[\"x\",2,3],[3,4,5]]") false ) )
  )
)

(deftest ut-getSSCellValues
  (testing "getSSCellValues(正常系)"
    (testing "セルアドレス指定1つ"
      (is
        (=
          (getSSCellValues
            "./resources/test01.xls"
            0
            "[[0,3]]")
          "[\"x\"]"
        )
      )
    )
    (testing "セルアドレス指定２つ"
      (is
        (=
          (getSSCellValues
            "./resources/test01.xls"
            0
            "[[0,3],[1,4]]")
          "[\"x\",\"p2\"]"
        )
      )
    )
    (testing "空のセルを指定"
      (is
        (=
          (getSSCellValues
            "./resources/test01.xls"
            0
            "[[100,100]]")
          "[\"\"]"
        )
      )
    )
    (testing "結果に空のセルを含む（最初）"
      (is
        (=
          (getSSCellValues
            "./resources/test01.xls"
            0
            "[[100,100],[0,1]]")
          "[\"\",1.0]"
        )
      )
    )
    (testing "結果に空のセルを含む（最後）"
      (is
        (=
          (getSSCellValues
            "./resources/test01.xls"
            0
            "[[1,1],[100,100]]")
          "[\"2013\\/07\\/29\",\"\"]"
        )
      )
    )
  )
)

; (deftest ut-getSSCellValues-xlsx
;   (testing "getSSCellValues(正常系)"
;     (testing "セルアドレス指定1つ"
;       (is
;         (=
;           (getSSCellValues
;             "./resources/test01.xlsx"
;             0
;             "[[0,3]]")
;           "[\"x\"]"
;         )
;       )
;     )
;     (testing "セルアドレス指定２つ"
;       (is
;         (=
;           (getSSCellValues
;             "./resources/test01.xlsx"
;             0
;             "[[0,3],[1,4]]")
;           "[\"x\",\"p2\"]"
;         )
;       )
;     )
;     (testing "空のセルを指定"
;       (is
;         (=
;           (getSSCellValues
;             "./resources/test01.xlsx"
;             0
;             "[[100,100]]")
;           "[\"\"]"
;         )
;       )
;     )
;     (testing "結果に空のセルを含む（最初）"
;       (is
;         (=
;           (getSSCellValues
;             "./resources/test01.xlsx"
;             0
;             "[[100,100],[0,1]]")
;           "[\"\",1.0]"
;         )
;       )
;     )
;     (testing "結果に空のセルを含む（最後）"
;       (is
;         (=
;           (getSSCellValues
;             "./resources/test01.xlsx"
;             0
;             "[[1,1],[100,100]]")
;           "[\"2013\\/07\\/29\",\"\"]"
;         )
;       )
;     )
;   )
; )

(deftest ut-getSSCellValues-ab
  (testing "getSSCellValues(異常系)"
    (testing "セルアドレス指定なし1"
      (is
        (=
          (try (getSSCellValues
            "./resources/test01.xls"
            0
            "")
          (catch Exception e (.getMessage e)))
          "Invalid cell address list. ()"
        )
      )
    )
    (testing "セルアドレス指定なし2"
      (is
        (=
          (try (getSSCellValues
            "./resources/test01.xls"
            0
            "[]")
          (catch Exception e (.getMessage e)))
          "Invalid cell address list. ([])"
        )
      )
    )
    (testing "エクセルファイルが存在しない"
      (is
        (=
          (try (getSSCellValues
            "X:/not-exists.file"
            0
            "[[0,3]]")
          (catch Exception e (.getName (class e))))
          "java.io.FileNotFoundException"
        )
      )
    )
    (testing "エクセルファイルを解析できない"
      (is
        (=
          (try (getSSCellValues
            "./resources/not-excel.txt"
            0
            "[[0,3]]")
          (catch Exception e (.getName (class e))))
          "java.lang.IllegalArgumentException"
        )
      )
    )
    (testing "エクセルファイルの指定のシートが存在しない"
      (is
        (=
          (try (getSSCellValues
            "./resources/test01.xls"
            10
            "[[0,3]]")
          (catch Exception e (vector (.getName (class e)) (.getMessage e))))
          ["java.lang.IllegalArgumentException" "Sheet index (10) is out of range (0..0)"]
        )
      )
    )
    (testing "セルアドレスリストを解析できない1"
      (is
        (=
          (try (getSSCellValues
            "./resources/test01.xls"
            0
            "")
          (catch Exception e (.getMessage e)))
          "Invalid cell address list. ()"  
        )
      )
    )
  )
)

; (deftest ut-getSSCellValues-ab-xlsx
;   (testing "getSSCellValues(異常系)"
;     (testing "セルアドレス指定なし1"
;       (is
;         (=
;           (try (getSSCellValues
;             "./resources/test01.xlsx"
;             0
;             "")
;           (catch Exception e (.getMessage e)))
;           "Invalid cell address list. ()"
;         )
;       )
;     )
;     (testing "セルアドレス指定なし2"
;       (is
;         (=
;           (try (getSSCellValues
;             "./resources/test01.xlsx"
;             0
;             "[]")
;           (catch Exception e (.getMessage e)))
;           "Invalid cell address list. ([])"
;         )
;       )
;     )
;     (testing "エクセルファイルが存在しない"
;       (is
;         (=
;           (try (getSSCellValues
;             "X:/not-exists.file"
;             0
;             "[[0,3]]")
;           (catch Exception e (.getName (class e))))
;           "java.io.FileNotFoundException"
;         )
;       )
;     )
;     (testing "エクセルファイルを解析できない"
;       (is
;         (=
;           (try (getSSCellValues
;             "./resources/not-excel.txt"
;             0
;             "[[0,3]]")
;           (catch Exception e (.getName (class e))))
;           "java.lang.IllegalArgumentException"
;         )
;       )
;     )
;     (testing "エクセルファイルの指定のシートが存在しない"
;       (is
;         (=
;           (try (getSSCellValues
;             "./resources/test01.xlsx"
;             10
;             "[[0,3]]")
;           (catch Exception e (vector (.getName (class e)) (.getMessage e))))
;           ["java.lang.IllegalArgumentException" "Sheet index (10) is out of range (0..0)"]
;         )
;       )
;     )
;     (testing "セルアドレスリストを解析できない1"
;       (is
;         (=
;           (try (getSSCellValues
;             "./resources/test01.xlsx"
;             0
;             "")
;           (catch Exception e (.getMessage e)))
;           "Invalid cell address list. ()"  
;         )
;       )
;     )
;   )
; )

(deftest ut-set-cell-value
  (testing "set-cell-value(正常系)"
    (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test05.xls")) 0)]
      (testing "文字列"
        (is
          (=
            (do
              (set-cell-value sheet 0 3 "x")
              (get-cell-value sheet 0 3)
            "x"
            )
          )
        )
      )
      (testing "文字列（マルチバイト）"
        (is
          (=
            (do
              (set-cell-value sheet 3 2 "有効：1／無効：0")  
              (get-cell-value sheet 3 2))
            "有効：1／無効：0"
          )
        )
      )
      (testing "空文字"
        (is
          (=
            (do
              (set-cell-value sheet 100 0 "")
              (get-cell-value sheet 100 0))
            ""
          )
        )
      )    
      (testing "数値"
        (is
          (==
            (do
              (set-cell-value sheet 0 1 1)  
              (get-cell-value sheet 0 1))
            1
          )
        )
      )   
      (testing "日付"
        (is
          (=
            (do
              (set-cell-value sheet 1 1 "2013/07/29")
              (get-cell-value sheet 1 1))
            "2013/07/29"
          )
        )
      )
      (testing "行が範囲外1"
        (is
          (=
            (do
              (set-cell-value sheet 0 -1 "x")
              (get-cell-value sheet 0 -1))
            ""
          )
        )
      )
      (testing "行が範囲外2"
        (is
          (=
            (do
              (set-cell-value sheet 0 65536 "x")
              (get-cell-value sheet 0 65536))
            ""
          )
        )
      )
      (testing "列が範囲外1"
        (is
          (=
            (do
              (set-cell-value sheet -1 3 "x")
              (get-cell-value sheet -1 3))
            ""
          )
        )
      )
      (testing "列が範囲外2"
        (is
          (=
            (do
              (set-cell-value sheet 256 3 "x")
              (get-cell-value sheet 256 3))
            ""
          )
        )
      )
    )
  )
)

; (deftest ut-set-cell-value-xlsx
;   (testing "set-cell-value(正常系)"
;     (let [sheet (.getSheetAt (WorkbookFactory/create (FileInputStream. "./resources/test05.xlsx")) 0)]
;       (testing "文字列"
;         (is
;           (=
;             (do
;               (set-cell-value sheet 0 3 "x")
;               (get-cell-value sheet 0 3)
;             "x"
;             )
;           )
;         )
;       )
;       (testing "文字列（マルチバイト）"
;         (is
;           (=
;             (do
;               (set-cell-value sheet 3 2 "有効：1／無効：0")  
;               (get-cell-value sheet 3 2))
;             "有効：1／無効：0"
;           )
;         )
;       )
;       (testing "空文字"
;         (is
;           (=
;             (do
;               (set-cell-value sheet 100 0 "")
;               (get-cell-value sheet 100 0))
;             ""
;           )
;         )
;       )    
;       (testing "数値"
;         (is
;           (==
;             (do
;               (set-cell-value sheet 0 1 1)  
;               (get-cell-value sheet 0 1))
;             1
;           )
;         )
;       )   
;       (testing "日付"
;         (is
;           (=
;             (do
;               (set-cell-value sheet 1 1 "2013/07/29")
;               (get-cell-value sheet 1 1))
;             "2013/07/29"
;           )
;         )
;       )
;       (testing "行が範囲外1"
;         (is
;           (=
;             (do
;               (set-cell-value sheet 0 -1 "x")
;               (get-cell-value sheet 0 -1))
;             ""
;           )
;         )
;       )
;       (testing "行が範囲外2"
;         (is
;           (=
;             (do
;               (set-cell-value sheet 0 65536 "x")
;               (get-cell-value sheet 0 65536))
;             ""
;           )
;         )
;       )
;       ; (testing "列が範囲外1"
;       ;   (is
;       ;     (=
;       ;       (do
;       ;         (set-cell-value sheet -1 3 "x")
;       ;         (get-cell-value sheet -1 3))
;       ;       ""
;       ;     )
;       ;   )
;       ; )
;       (testing "列が範囲外2"
;         (is
;           (=
;             (do
;               (set-cell-value sheet 16384 3 "x")
;               (get-cell-value sheet 16384 3))
;             ""
;           )
;         )
;       )
;     )
;   )
; )

(deftest ut-setSSCellValues
  (testing "setSSCellValues(正常系)"
    (testing "セルアドレス指定1つ"
      (is
        (=
          (do
            (setSSCellValues
              "./resources/test06.xls"
              0
              "[[0,3,\"x\"]]")
            (getSSCellValues
              "./resources/test06.xls"
              0
              "[[0,3]]"))
          "[\"x\"]"
        )
      )
    )
    (testing "セルアドレス指定２つ"
      (is
        (=
          (do
            (setSSCellValues
              "./resources/test06.xls"
              0
              "[[0,3,\"x\"],[1,4,\"p2\"]]")
            (getSSCellValues
              "./resources/test06.xls"
              0
              "[[0,3],[1,4]]"))
          "[\"x\",\"p2\"]"
        )
      )
    )
    (testing "空指定"
      (is
        (=
          (do
            (setSSCellValues
              "./resources/test06.xls"
              0
              "[[100,100,\"\"]]")
            (getSSCellValues
              "./resources/test06.xls"
              0
              "[[100,100]]"))
          "[\"\"]"
        )
      )
    )
  )
)

; (deftest ut-setSSCellValues-xlsx
;   (testing "setSSCellValues(正常系)"
;     (testing "セルアドレス指定1つ"
;       (is
;         (=
;           (do
;             (setSSCellValues
;               "./resources/test06.xlsx"
;               0
;               "[[0,3,\"x\"]]")
;             (getSSCellValues
;               "./resources/test06.xlsx"
;               0
;               "[[0,3]]"))
;           "[\"x\"]"
;         )
;       )
;     )
;     (testing "セルアドレス指定２つ"
;       (is
;         (=
;           (do
;             (setSSCellValues
;               "./resources/test06.xlsx"
;               0
;               "[[0,3,\"x\"],[1,4,\"p2\"]]")
;             (getSSCellValues
;               "./resources/test06.xlsx"
;               0
;               "[[0,3],[1,4]]"))
;           "[\"x\",\"p2\"]"
;         )
;       )
;     )
;     (testing "空指定"
;       (is
;         (=
;           (do
;             (setSSCellValues
;               "./resources/test06.xlsx"
;               0
;               "[[100,100,\"\"]]")
;             (getSSCellValues
;               "./resources/test06.xlsx"
;               0
;               "[[100,100]]"))
;           "[\"\"]"
;         )
;       )
;     )
;   )
; )

(deftest ut-setSSCellValues-ab
  (testing "setSSCellValues(異常系)"
    (testing "指定なし1"
      (is
        (=
          (try (setSSCellValues
            "./resources/test06.xls"
            0
            "")
          (catch Exception e (.getMessage e)))
          "Invalid cell address and value list. ()"
        )
      )
    )
    (testing "指定なし2"
      (is
        (=
          (try (setSSCellValues
            "./resources/test06.xls"
            0
            "[]")
          (catch Exception e (.getMessage e)))
          "Invalid cell address and value list. ([])"
        )
      )
    )
    (testing "エクセルファイルが存在しない"
      (is
        (=
          (try (setSSCellValues
            "X:/not-exists.file"
            0
            "[[0,3,\"x\"]]")
          (catch Exception e (.getName (class e))))
          "java.io.FileNotFoundException"
        )
      )
    )
    (testing "エクセルファイルを解析できない"
      (is
        (=
          (try (setSSCellValues
            "./resources/not-excel.txt"
            0
            "[[0,3,\"x\"]]")
          (catch Exception e (.getName (class e))))
          "java.lang.IllegalArgumentException"
        )
      )
    )
    (testing "エクセルファイルの指定のシートが存在しない"
      (is
        (=
          (try (setSSCellValues
            "./resources/test01.xls"
            10
            "[[0,3,\"x\"]]")
          (catch Exception e (vector (.getName (class e)) (.getMessage e))))
          ["java.lang.IllegalArgumentException" "Sheet index (10) is out of range (0..0)"]
        )
      )
    )
  )
)

; (deftest ut-setSSCellValues-ab-xlsx
;   (testing "setSSCellValues(異常系)"
;     (testing "指定なし1"
;       (is
;         (=
;           (try (setSSCellValues
;             "./resources/test06.xlsx"
;             0
;             "")
;           (catch Exception e (.getMessage e)))
;           "Invalid cell address and value list. ()"
;         )
;       )
;     )
;     (testing "指定なし2"
;       (is
;         (=
;           (try (setSSCellValues
;             "./resources/test06.xlsx"
;             0
;             "[]")
;           (catch Exception e (.getMessage e)))
;           "Invalid cell address and value list. ([])"
;         )
;       )
;     )
;     (testing "エクセルファイルが存在しない"
;       (is
;         (=
;           (try (setSSCellValues
;             "X:/not-exists.file"
;             0
;             "[[0,3,\"x\"]]")
;           (catch Exception e (.getName (class e))))
;           "java.io.FileNotFoundException"
;         )
;       )
;     )
;     (testing "エクセルファイルを解析できない"
;       (is
;         (=
;           (try (setSSCellValues
;             "./resources/not-excel.txt"
;             0
;             "[[0,3,\"x\"]]")
;           (catch Exception e (.getName (class e))))
;           "java.lang.IllegalArgumentException"
;         )
;       )
;     )
;     (testing "エクセルファイルの指定のシートが存在しない"
;       (is
;         (=
;           (try (setSSCellValues
;             "./resources/test01.xlsx"
;             10
;             "[[0,3,\"x\"]]")
;           (catch Exception e (vector (.getName (class e)) (.getMessage e))))
;           ["java.lang.IllegalArgumentException" "Sheet index (10) is out of range (0..0)"]
;         )
;       )
;     )
;   )
; )

(deftest ut-insertSS
  (testing "insertSS(正常系)"
    (testing "2属性、必須属性完備"
      (copy (file "./resources/test07.xls") (file "./resources/work_test07.xls"))
      (is
        (=
          (do
            (insertSS
              "./resources/test01.json"
              "./resources/work_test07.xls"
              "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
            (selectSS
              "./resources/test01.json"
              "./resources/work_test07.xls"
              "{ \"attributes\" : [\"id\", \"pwd\"] }")
          )
          "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"x\",\"pwd\":\"p1\"}]"
        )
      )
      (delete-file "./resources/work_test07.xls")
    )
    (testing "2属性、必須属性完備(既存データ1件あり1)"
      (copy (file "./resources/test08.xls") (file "./resources/work_test08.xls"))
      (is
        (=
          (do
            (insertSS
              "./resources/test01.json"
              "./resources/work_test08.xls"
              "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
            (selectSS
              "./resources/test01.json"
              "./resources/work_test08.xls"
              "{ \"attributes\" : [\"id\", \"pwd\"] }")
          )
          "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"z\",\"pwd\":\"z\"},{\"id\":\"x\",\"pwd\":\"p1\"}]"
        )
      )
      (delete-file "./resources/work_test08.xls")
    )
    (testing "2属性、必須属性完備(既存データ1件あり2)"
      (copy (file "./resources/test09.xls") (file "./resources/work_test09.xls"))
      (is
        (=
          (do
            (insertSS
              "./resources/test01.json"
              "./resources/work_test09.xls"
              "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
            (selectSS
              "./resources/test01.json"
              "./resources/work_test09.xls"
              "{ \"attributes\" : [\"id\", \"pwd\"] }")
          )
          "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"z\",\"pwd\":\"z\"},{\"id\":\"x\",\"pwd\":\"p1\"}]"
        )
      )
      (delete-file "./resources/work_test09.xls")
    )
    (testing "2属性、必須属性完備(既存データ1件あり3)"
      (copy (file "./resources/test10.xls") (file "./resources/work_test10.xls"))
      (is
        (=
          (do
            (insertSS
              "./resources/test01.json"
              "./resources/work_test10.xls"
              "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
            (selectSS
              "./resources/test01.json"
              "./resources/work_test10.xls"
              "{ \"attributes\" : [\"id\", \"pwd\"] }")
          )
          "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"z\",\"pwd\":\"z\"},{\"id\":\"x\",\"pwd\":\"p1\"}]"
        )
      )
      (delete-file "./resources/work_test10.xls")
    )
  )
)

; (deftest ut-insertSS-xlsx
;   (testing "insertSS(正常系)"
;     (testing "2属性、必須属性完備"
;       (copy (file "./resources/test07.xlsx") (file "./resources/work_test07.xlsx"))
;       (is
;         (=
;           (do
;             (insertSS
;               "./resources/test01.json"
;               "./resources/work_test07.xlsx"
;               "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
;             (selectSS
;               "./resources/test01.json"
;               "./resources/work_test07.xlsx"
;               "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           )
;           "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"x\",\"pwd\":\"p1\"}]"
;         )
;       )
;       (delete-file "./resources/work_test07.xlsx")
;     )
;     (testing "2属性、必須属性完備(既存データ1件あり1)"
;       (copy (file "./resources/test08.xlsx") (file "./resources/work_test08.xlsx"))
;       (is
;         (=
;           (do
;             (insertSS
;               "./resources/test01.json"
;               "./resources/work_test08.xlsx"
;               "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
;             (selectSS
;               "./resources/test01.json"
;               "./resources/work_test08.xlsx"
;               "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           )
;           "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"z\",\"pwd\":\"z\"},{\"id\":\"x\",\"pwd\":\"p1\"}]"
;         )
;       )
;       (delete-file "./resources/work_test08.xlsx")
;     )
;     (testing "2属性、必須属性完備(既存データ1件あり2)"
;       (copy (file "./resources/test09.xlsx") (file "./resources/work_test09.xlsx"))
;       (is
;         (=
;           (do
;             (insertSS
;               "./resources/test01.json"
;               "./resources/work_test09.xlsx"
;               "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
;             (selectSS
;               "./resources/test01.json"
;               "./resources/work_test09.xlsx"
;               "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           )
;           "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"z\",\"pwd\":\"z\"},{\"id\":\"x\",\"pwd\":\"p1\"}]"
;         )
;       )
;       (delete-file "./resources/work_test09.xlsx")
;     )
;     (testing "2属性、必須属性完備(既存データ1件あり3)"
;       (copy (file "./resources/test10.xlsx") (file "./resources/work_test10.xlsx"))
;       (is
;         (=
;           (do
;             (insertSS
;               "./resources/test01.json"
;               "./resources/work_test10.xlsx"
;               "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
;             (selectSS
;               "./resources/test01.json"
;               "./resources/work_test10.xlsx"
;               "{ \"attributes\" : [\"id\", \"pwd\"] }")
;           )
;           "[{\"id\":\"y\",\"pwd\":\"p2\"},{\"id\":\"z\",\"pwd\":\"z\"},{\"id\":\"x\",\"pwd\":\"p1\"}]"
;         )
;       )
;       (delete-file "./resources/work_test10.xlsx")
;     )
;   )
; )

(deftest ut-insertSS-ab
  (testing "insertSS(異常系)"
    (testing "存在しない属性を指定"
      (is
        (let [schema-file-name-name "./resources/test01.json"]
          (=
            (try
              (insertSS
                schema-file-name-name
                "./resources/test11.xls"
                "[ { \"id\" : \"x\", \"pwd\" : \"p1\", \"notexistattr\" \"p1\"}, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
              (catch Exception e (.getMessage e)))
            (str "Record ({:id \"x\", :pwd \"p1\", :notexistattr \"p1\"}) is not consistent with schema definition in the file (" schema-file-name-name ").")
          )
        )
      )
    )
    (testing "必須属性を未指定"
      (is
        (let [schema-file-name-name "./resources/test01.json"]
          (=
            (try
              (insertSS
                schema-file-name-name
                "./resources/test11.xls"
                "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"pwd\" : \"p2\" }]")
              (catch Exception e (.getMessage e)))
            (str "Record ({:pwd \"p2\"}) is not consistent with schema definition in the file (" schema-file-name-name ").")
          )
        )
      )
    )
    (testing "空きの行が足りない"
      (is
        (=
          (try
            (insertSS
              "./resources/test01.json"
              "./resources/test13.xls"
              "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
            (catch Exception e (.getMessage e)))
          "2 rows insert failed (all row). Available row count is 0."
        )
      )
    )
  )
)

; (deftest ut-insertSS-ab-xlsx
;   (testing "insertSS(異常系)"
;     (testing "存在しない属性を指定"
;       (is
;         (let [schema-file-name-name "./resources/test01.json"]
;           (=
;             (try
;               (insertSS
;                 schema-file-name-name
;                 "./resources/test11.xlsx"
;                 "[ { \"id\" : \"x\", \"pwd\" : \"p1\", \"notexistattr\" \"p1\"}, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
;               (catch Exception e (.getMessage e)))
;             (str "Record ({:id \"x\", :pwd \"p1\", :notexistattr \"p1\"}) is not consistent with schema definition in the file (" schema-file-name-name ").")
;           )
;         )
;       )
;     )
;     (testing "必須属性を未指定"
;       (is
;         (let [schema-file-name-name "./resources/test01.json"]
;           (=
;             (try
;               (insertSS
;                 schema-file-name-name
;                 "./resources/test11.xlsx"
;                 "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"pwd\" : \"p2\" }]")
;               (catch Exception e (.getMessage e)))
;             (str "Record ({:pwd \"p2\"}) is not consistent with schema definition in the file (" schema-file-name-name ").")
;           )
;         )
;       )
;     )
;     (testing "空きの行が足りない"
;       (is
;         (=
;           (try
;             (insertSS
;               "./resources/test01.json"
;               "./resources/test13.xlsx"
;               "[ { \"id\" : \"x\", \"pwd\" : \"p1\" }, { \"id\" : \"y\", \"pwd\" : \"p2\" }]")
;             (catch Exception e (.getMessage e)))
;           "2 rows insert failed (all row). Available row count is 0."
;         )
;       )
;     )
;   )
; )

(deftest ut-updateSS
  (testing "updateSS(正常系)"
    (testing "1属性指定、1属性更新、1レコードずつに影響"
      (Thread/sleep 1000)
      (copy (file "./resources/test12.xls") (file "./resources/work_test12.xls"))
      (is
        (= 
          (do
            (updateSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "[ { \"pwd\" : \"p11\", \"whereClause\" : { \"id\" : \"x\" } }
               , { \"pwd\" : \"p22\", \"whereClause\" : { \"id\" : \"y\" } } ]")
            (selectSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
          )
          "[{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p22\"},{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p11\"}]"
        )
      )
      (delete-file "./resources/work_test12.xls")
    )
    (testing "1属性指定、2属性更新、1レコードに影響"
      (Thread/sleep 1000)
      (copy (file "./resources/test12.xls") (file "./resources/work_test12.xls"))
      (is
        (= 
          (do
            (updateSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "[ { \"pwd\" : \"p11\", \"valid_flg\" : \"1\", \"whereClause\" : { \"id\" : \"x\" } } ]")
            (selectSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\", \"valid_flg\"] }")
          )
          "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p11\",\"valid_flg\":\"1\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\",\"valid_flg\":\"\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\",\"valid_flg\":\"\"}]"
        )
      )
      (delete-file "./resources/work_test12.xls")
    )
    (testing "1属性指定、1属性更新、1レコードずつに影響（値にマルチバイト文字）"
      (Thread/sleep 1000)
      (copy (file "./resources/test12.xls") (file "./resources/work_test12.xls"))
      (is
        (= 
          (do
            (updateSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "[ { \"pwd\" : \"パスワードイチ号ｘＸ\", \"whereClause\" : { \"id\" : \"x\" } } ]")
            (selectSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
          )
          "[{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"},{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"\\u30d1\\u30b9\\u30ef\\u30fc\\u30c9\\u30a4\\u30c1\\u53f7\\uff58\\uff38\"}]"
        )
      )
      (delete-file "./resources/work_test12.xls")
    )
    (testing "1属性指定、1属性更新、1レコードずつに影響（条件にマルチバイト文字）"
      (Thread/sleep 1000)
      (copy (file "./resources/test12.xls") (file "./resources/work_test12.xls"))
      (is
        (= 
          (do
            (updateSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "[ { \"pwd\" : \"ppppp\", \"whereClause\" : { \"ref_num\" : \"ａＡあア亜\" } } ]")
            (selectSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
          )
          "[{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"},{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"ppppp\"}]"
        )
      )
      (delete-file "./resources/work_test12.xls")
    )
    (testing "1属性指定、1属性更新、1レコードずつに影響（値に空文字）"
      (Thread/sleep 1000)
      (copy (file "./resources/test12.xls") (file "./resources/work_test12.xls"))
      (is
        (= 
          (do
            (updateSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "[ { \"pwd\" : \"\", \"whereClause\" : { \"id\" : \"x\" } } ]")
            (selectSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
          )
          "[{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"}]"
        )
      )
      (delete-file "./resources/work_test12.xls")
    )
    (testing "1属性指定、1属性更新、1レコードずつに影響（条件に空文字）"
      (Thread/sleep 1000)
      (copy (file "./resources/test12.xls") (file "./resources/work_test12.xls"))
      (is
        (= 
          (do
            (updateSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "[ { \"pwd\" : \"pwdxxx\", \"whereClause\" : { \"ref_num\" : \"\" } } ]")
            (selectSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
          )
          "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p1\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"pwdxxx\"}]"
        )
      )
      (delete-file "./resources/work_test12.xls")
    )
    (testing "1属性指定、1属性更新、2レコードに影響"
      (Thread/sleep 1000)
      (copy (file "./resources/test12.xls") (file "./resources/work_test12.xls"))
      (is
        (=
          (do
            (updateSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "[ { \"pwd\" : \"p111\", \"whereClause\" : { \"host\" : \"h1\" } } ]")
            (selectSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
          )
          "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p111\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p111\"}]"
        )
      )
      (delete-file "./resources/work_test12.xls")
    )
    (testing "2属性指定、1属性更新、1レコードに影響"
      (Thread/sleep 1000)
      (copy (file "./resources/test12.xls") (file "./resources/work_test12.xls"))
      (is
        (=
          (do
            (updateSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "[ { \"pwd\" : \"p222\", \"whereClause\" : { \"host\" : \"h1\", \"id\" : \"y\" } } ]")
            (selectSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
          )
          "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p1\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p222\"}]"
        )
      )
      (delete-file "./resources/work_test12.xls")
    )
    (testing "2属性指定、1属性更新、0レコードに影響(ヒットするレコードなし)"
      (Thread/sleep 1000)
      (copy (file "./resources/test12.xls") (file "./resources/work_test12.xls"))
      (is
        (=
          (do
            (updateSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "[ { \"pwd\" : \"xxxxxxx\", \"whereClause\" : { \"host\" : \"h2\", \"id\" : \"y\" } } ]")
            (selectSS
              "./resources/test08.json"
              "./resources/work_test12.xls"
              "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
          )
          "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p1\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"}]"
        )
      )
      (delete-file "./resources/work_test12.xls")
    )
  )
)

; (deftest ut-updateSS-xlsx
;   (testing "updateSS(正常系)"
;     (testing "1属性指定、1属性更新、1レコードずつに影響"
;       (Thread/sleep 1000)
;       (copy (file "./resources/test12.xlsx") (file "./resources/work_test12.xlsx"))
;       (is
;         (= 
;           (do
;             (updateSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "[ { \"pwd\" : \"p11\", \"whereClause\" : { \"id\" : \"x\" } }
;                , { \"pwd\" : \"p22\", \"whereClause\" : { \"id\" : \"y\" } } ]")
;             (selectSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
;           )
;           "[{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p22\"},{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p11\"}]"
;         )
;       )
;       (delete-file "./resources/work_test12.xlsx")
;     )
;     (testing "1属性指定、2属性更新、1レコードに影響"
;       (Thread/sleep 1000)
;       (copy (file "./resources/test12.xlsx") (file "./resources/work_test12.xlsx"))
;       (is
;         (= 
;           (do
;             (updateSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "[ { \"pwd\" : \"p11\", \"valid_flg\" : \"1\", \"whereClause\" : { \"id\" : \"x\" } } ]")
;             (selectSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\", \"valid_flg\"] }")
;           )
;           "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p11\",\"valid_flg\":\"1\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\",\"valid_flg\":\"\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\",\"valid_flg\":\"\"}]"
;         )
;       )
;       (delete-file "./resources/work_test12.xlsx")
;     )
;     (testing "1属性指定、1属性更新、1レコードずつに影響（値にマルチバイト文字）"
;       (Thread/sleep 1000)
;       (copy (file "./resources/test12.xlsx") (file "./resources/work_test12.xlsx"))
;       (is
;         (= 
;           (do
;             (updateSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "[ { \"pwd\" : \"パスワードイチ号ｘＸ\", \"whereClause\" : { \"id\" : \"x\" } } ]")
;             (selectSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
;           )
;           "[{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"},{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"\\u30d1\\u30b9\\u30ef\\u30fc\\u30c9\\u30a4\\u30c1\\u53f7\\uff58\\uff38\"}]"
;         )
;       )
;       (delete-file "./resources/work_test12.xlsx")
;     )
;     (testing "1属性指定、1属性更新、1レコードずつに影響（条件にマルチバイト文字）"
;       (Thread/sleep 1000)
;       (copy (file "./resources/test12.xlsx") (file "./resources/work_test12.xlsx"))
;       (is
;         (= 
;           (do
;             (updateSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "[ { \"pwd\" : \"ppppp\", \"whereClause\" : { \"ref_num\" : \"ａＡあア亜\" } } ]")
;             (selectSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
;           )
;           "[{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"},{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"ppppp\"}]"
;         )
;       )
;       (delete-file "./resources/work_test12.xlsx")
;     )
;     (testing "1属性指定、1属性更新、1レコードずつに影響（値に空文字）"
;       (Thread/sleep 1000)
;       (copy (file "./resources/test12.xlsx") (file "./resources/work_test12.xlsx"))
;       (is
;         (= 
;           (do
;             (updateSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "[ { \"pwd\" : \"\", \"whereClause\" : { \"id\" : \"x\" } } ]")
;             (selectSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
;           )
;           "[{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"}]"
;         )
;       )
;       (delete-file "./resources/work_test12.xlsx")
;     )
;     (testing "1属性指定、1属性更新、1レコードずつに影響（条件に空文字）"
;       (Thread/sleep 1000)
;       (copy (file "./resources/test12.xlsx") (file "./resources/work_test12.xlsx"))
;       (is
;         (= 
;           (do
;             (updateSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "[ { \"pwd\" : \"pwdxxx\", \"whereClause\" : { \"ref_num\" : \"\" } } ]")
;             (selectSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
;           )
;           "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p1\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"pwdxxx\"}]"
;         )
;       )
;       (delete-file "./resources/work_test12.xlsx")
;     )
;     (testing "1属性指定、1属性更新、2レコードに影響"
;       (Thread/sleep 1000)
;       (copy (file "./resources/test12.xlsx") (file "./resources/work_test12.xlsx"))
;       (is
;         (=
;           (do
;             (updateSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "[ { \"pwd\" : \"p111\", \"whereClause\" : { \"host\" : \"h1\" } } ]")
;             (selectSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
;           )
;           "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p111\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p111\"}]"
;         )
;       )
;       (delete-file "./resources/work_test12.xlsx")
;     )
;     (testing "2属性指定、1属性更新、1レコードに影響"
;       (Thread/sleep 1000)
;       (copy (file "./resources/test12.xlsx") (file "./resources/work_test12.xlsx"))
;       (is
;         (=
;           (do
;             (updateSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "[ { \"pwd\" : \"p222\", \"whereClause\" : { \"host\" : \"h1\", \"id\" : \"y\" } } ]")
;             (selectSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
;           )
;           "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p1\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p222\"}]"
;         )
;       )
;       (delete-file "./resources/work_test12.xlsx")
;     )
;     (testing "2属性指定、1属性更新、0レコードに影響(ヒットするレコードなし)"
;       (Thread/sleep 1000)
;       (copy (file "./resources/test12.xlsx") (file "./resources/work_test12.xlsx"))
;       (is
;         (=
;           (do
;             (updateSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "[ { \"pwd\" : \"xxxxxxx\", \"whereClause\" : { \"host\" : \"h2\", \"id\" : \"y\" } } ]")
;             (selectSS
;               "./resources/test08.json"
;               "./resources/work_test12.xlsx"
;               "{ \"attributes\" : [\"host\", \"id\", \"pwd\"] }")
;           )
;           "[{\"host\":\"h1\",\"id\":\"x\",\"pwd\":\"p1\"},{\"host\":\"h2\",\"id\":\"z\",\"pwd\":\"p3\"},{\"host\":\"h1\",\"id\":\"y\",\"pwd\":\"p2\"}]"
;         )
;       )
;       (delete-file "./resources/work_test12.xlsx")
;     )
;   )
; )

(deftest ut-updateSS-ab
  (testing "updateSS(異常系)"
    (testing "存在しない属性を値に指定"
      (is
        (let [schema-file-name-name "./resources/test08.json"]
          (=
            (try
              (updateSS
                schema-file-name-name
                "./resources/test12.xls"
                "[ { \"pwd\" : \"p11\", \"whereClause\" : { \"id\" : \"x\" } }
                 , { \"notexistattr\" : \"p22\", \"whereClause\" : { \"id\" : \"y\" } } ]")
              (catch Exception e (.getMessage e)))
            (str "Record ({:notexistattr \"p22\", :whereClause {:id \"y\"}}) is not consistent with schema definition in the file (" schema-file-name-name ").")
          )
        )
      )
    )
    (testing "存在しない属性をwhere句に指定"
      (is
        (let [schema-file-name-name "./resources/test08.json"]
          (=
            (try
              (updateSS
                schema-file-name-name
                "./resources/test12.xls"
                "[ { \"pwd\" : \"p11\", \"whereClause\" : { \"id\" : \"x\" } }
                 , { \"pwd\" : \"p22\", \"whereClause\" : { \"notexistattr\" : \"y\" } } ]")
              (catch Exception e (.getMessage e)))
            (str "Record ({:pwd \"p22\", :whereClause {:notexistattr \"y\"}}) is not consistent with schema definition in the file (" schema-file-name-name ").")
          )
        )
      )
    )
    (testing "必須属性を空に更新"
      (is
        (let [schema-file-name-name "./resources/test08.json"]
          (=
            (try
              (updateSS
                schema-file-name-name
                "./resources/test12.xls"
                "[ { \"pwd\" : \"p11\", \"whereClause\" : { \"id\" : \"x\" } }
                 , { \"host\" : \"\", \"whereClause\" : { \"id\" : \"y\" } } ]")
              (catch Exception e (.getMessage e)))
            (str "Record ({:host \"\", :whereClause {:id \"y\"}}) is not consistent with schema definition in the file (" schema-file-name-name ").")
          )
        )
      )
    )
  )
)

; (deftest ut-updateSS-ab-xlsx
;   (testing "updateSS(異常系)"
;     (testing "存在しない属性を値に指定"
;       (is
;         (let [schema-file-name-name "./resources/test08.json"]
;           (=
;             (try
;               (updateSS
;                 schema-file-name-name
;                 "./resources/test12.xlsx"
;                 "[ { \"pwd\" : \"p11\", \"whereClause\" : { \"id\" : \"x\" } }
;                  , { \"notexistattr\" : \"p22\", \"whereClause\" : { \"id\" : \"y\" } } ]")
;               (catch Exception e (.getMessage e)))
;             (str "Record ({:notexistattr \"p22\", :whereClause {:id \"y\"}}) is not consistent with schema definition in the file (" schema-file-name-name ").")
;           )
;         )
;       )
;     )
;     (testing "存在しない属性をwhere句に指定"
;       (is
;         (let [schema-file-name-name "./resources/test08.json"]
;           (=
;             (try
;               (updateSS
;                 schema-file-name-name
;                 "./resources/test12.xlsx"
;                 "[ { \"pwd\" : \"p11\", \"whereClause\" : { \"id\" : \"x\" } }
;                  , { \"pwd\" : \"p22\", \"whereClause\" : { \"notexistattr\" : \"y\" } } ]")
;               (catch Exception e (.getMessage e)))
;             (str "Record ({:pwd \"p22\", :whereClause {:notexistattr \"y\"}}) is not consistent with schema definition in the file (" schema-file-name-name ").")
;           )
;         )
;       )
;     )
;     (testing "必須属性を空に更新"
;       (is
;         (let [schema-file-name-name "./resources/test08.json"]
;           (=
;             (try
;               (updateSS
;                 schema-file-name-name
;                 "./resources/test12.xlsx"
;                 "[ { \"pwd\" : \"p11\", \"whereClause\" : { \"id\" : \"x\" } }
;                  , { \"host\" : \"\", \"whereClause\" : { \"id\" : \"y\" } } ]")
;               (catch Exception e (.getMessage e)))
;             (str "Record ({:host \"\", :whereClause {:id \"y\"}}) is not consistent with schema definition in the file (" schema-file-name-name ").")
;           )
;         )
;       )
;     )
;   )
; )
