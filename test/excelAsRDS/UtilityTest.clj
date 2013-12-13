(ns excelAsRDS.UtilityTest
  (:require
    [clojure.test :refer :all]
    [excelAsRDS.Utility :refer :all]))

(deftest ut-selectDB
  (testing "0要素(オブジェクト)が等しい"
    (is
      (=
        (selectDB
          "com.microsoft.jdbc.sqlserver.SQLServerDriver"
          "sqlserver"
          "//spam13:1433;database=mdb;user=ServiceDesk;password=Passw0rd"
          "SELECT * FROM ca_contact")
        "[{\"val\":\"0\"},{\"val\":\"1\"},{\"val\":\"2\"},{\"val\":\"3\"}]"
      )
    )
  )
)

(deftest ut-isEqualJSONStrAsSet
  (testing "isEqualJSONStrAsSet(正常系)"
    (testing "0要素(オブジェクト)が等しい"
      (is
        (isEqualJSONStrAsSet
          "[]"
          "[]"
        )
      )
    )
    (testing "1要素(リテラル)が等しい"
      (is
        (isEqualJSONStrAsSet
          "[1]"
          "[1]"
        )
      )
    )
    (testing "1要素(リテラル)が等しくない1"
      (is
        (not (isEqualJSONStrAsSet
          "[1]"
          "[2]"
        ))
      )
    )
    (testing "1要素(リテラル)が等しくない2"
      (is
        (not (isEqualJSONStrAsSet
          "[1]"
          "[]"
        ))
      )
    )
    (testing "2要素(リテラル)が等しい"
      (is
        (isEqualJSONStrAsSet
          "[1 2]"
          "[1 2]"
        )
      )
    )
    (testing "2要素(リテラル)が等しくない1"
      (is
        (not (isEqualJSONStrAsSet
          "[1 2]"
          "[1 3]"
        ))
      )
    )
    (testing "2要素(リテラル)が等しくない2"
      (is
        (not (isEqualJSONStrAsSet
          "[2 2]"
          "[1 2]"
        ))
      )
    )
    (testing "1要素(オブジェクト)が等しい1"
      (is
        (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}]"
          "[{\"id\" : \"x\"}]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しい2"
      (is
        (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}]"
          "[{\"id\" : \"x\"}, {\"id\" : \"x\"}]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しくない1"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}]"
          "[{\"id\" : \"y\"}]"
        ))
      )
    )
    (testing "1要素(オブジェクト)が等しくない2"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"pwd\" : \"x\"}]"
          "[{\"id\" : \"x\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しい1"
      (is
        (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しい2"
      (is
        (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"pwd\" : \"px\"}, {\"id\" : \"x\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない1"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"id\" : \"y\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しくない2"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"identity\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しくない3"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"py\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しくない4"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"host\" : \"hx\"}, {\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しくない5"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}, {\"host\" : \"hx\"}]"
        ))
      )
    )
    (testing "0要素(オブジェクト)が等しい"
      (is
        (isEqualJSONStrAsSet
          "[]"
          "[]"
        )
      )
    )
  )
)

(deftest ut-differenceJSONStrAsSet
  (testing "differenceJSONStrAsSet(正常系)"
    (testing "0要素(オブジェクト)が等しい"
      (is
        (=
          (differenceJSONStrAsSet
            "[]"
            "[]"
          )
          "[]"
        )
      )
    )
    (testing "1要素(リテラル)が等しい"
      (is
        (=
          (differenceJSONStrAsSet
            "[1]"
            "[1]"
          )
        )
        "[]"
      )
    )
    (testing "1要素(リテラル)が等しくない1"
      (is
        (=
          (differenceJSONStrAsSet
            "[1]"
            "[2]"
          )
          "[1]"
        )
      )
    )
    (testing "1要素(リテラル)が等しくない2"
      (is
        (=
          (differenceJSONStrAsSet
            "[1]"
            "[]"
          )
          "[1]"
        )
      )
    )
    (testing "2要素(リテラル)が等しい"
      (is
        (=
          (differenceJSONStrAsSet
            "[1 2]"
            "[1 2]"
          )
          "[]"
        )
      )
    )
    (testing "2要素(リテラル)が等しくない1"
      (is
        (=
          (differenceJSONStrAsSet
            "[1 2]"
            "[1 3]"
          )
          "[2]"
        )
      )
    )
    (testing "2要素(リテラル)が等しくない2"
      (is
        (=
          (differenceJSONStrAsSet
            "[2 2]"
            "[1 2]"
          )
          "[]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しい1"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}]"
            "[{\"id\" : \"x\"}]"
          )
          "[]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しい2"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}]"
            "[{\"id\" : \"x\"}, {\"id\" : \"x\"}]"
          )
          "[]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しくない1"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}]"
            "[{\"id\" : \"y\"}]"
          )
          "[{\"id\":\"x\"}]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しくない2"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"pwd\" : \"x\"}]"
            "[{\"id\" : \"x\"}]"
          )
          "[{\"pwd\":\"x\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しい1"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          )
          "[]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しい2"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"pwd\" : \"px\"}, {\"id\" : \"x\"}]"
          )
          "[]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない1"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"y\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          )
          "[{\"id\":\"y\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない2"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"identity\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          )
          "[{\"identity\":\"x\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない3"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"py\"}]"
          )
          "[{\"pwd\":\"px\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない4"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"host\" : \"hx\"}, {\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          )
          "[{\"host\":\"hx\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない5"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}, {\"host\" : \"hx\"}]"
          )
          "[]"
        )
      )
    )
    (testing "0要素(オブジェクト)が等しい"
      (is
        (=
          (differenceJSONStrAsSet
            "[]"
            "[]"
          )
          "[]"
        )
      )
    )
  )
)