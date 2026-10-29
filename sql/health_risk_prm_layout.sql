/* =====================================================================
   건강보험 위험보험료/손해액 집계 (레이아웃.xlsx 기준) - 초안
   - 기준 테이블 : DHADM.STAT_RISK_PRM_MASTER (A)
   - 고객역할 테이블 (G) : EAGLE.GOGEKYUKHAL  증권번호(BHGY_JGNO) → 고객PK(GG_PK), 피보험자 역할 GGYH_CODE = '21'
   - 고객 테이블     (H) : EAGLE.GOGEK        고객PK(PK) → 피보험자주민번호(GGNO)
   - 유의설계사 테이블 (C) : YUI_SAWONCD 보유 테이블       ※ 실제 테이블/컬럼명으로 교체
   - 범위 : 건강보험 한정 SUBSTR(MAIN_PROD_CODE,1,2) = '4A'
   ※ GROUP BY 절에는 서브쿼리(IN (SELECT ...))를 쓸 수 없으므로(ORA-22818)
     인라인뷰(BASE)에서 구분값을 먼저 만들고 바깥에서 집계한다.
   ※ 본인계약/유의설계사 여부는 EXISTS 로 판정 → H/C 테이블 건수와 무관하게 A 행이 복제되지 않음
   ※ 건강보험은 종피 없이 주피 1인 → 피보험자 구분 조건 불필요
   ===================================================================== */
SELECT  BASE.계약년월
      , BASE.마감년월
      , BASE.연령대
      , BASE.담보코드
      , BASE.심사유형
      , BASE.모집지점명
      , BASE.본인계약여부
      , BASE.유의설계사여부
      , SUM(BASE.RISK_PRM)  AS 위험보험료합
      , SUM(BASE.LOSS_AMT)  AS 손해액합
  FROM (
        SELECT  SUBSTR(A.POLI_DT, 1, 6)                       AS 계약년월
              , A.CLOS_YR_MON                                 AS 마감년월
              -- *연령대 : 10세 단위 (피보험자나이 = 가입연령 AGE, 도달연령 기준이면 ATTAIN_AGE 로 교체)
              , NVL2(A.AGE
                   , LPAD(FLOOR(A.AGE / 10) * 10, 2, '0') || '~' ||
                     LPAD(FLOOR(A.AGE / 10) * 10 + 9, 2, '0')
                   , '미상')                                  AS 연령대
              , A.GRNTE_CODE                                  AS 담보코드   -- 신위험보장급부코드
              , A.DISCLOSURE_TYPE                             AS 심사유형   -- 일반심사/간편305/간편335/간편355/간편3105
              , A.RCRT_BRNC                                   AS 모집지점명
              -- **본인계약여부 : 모집사원실명번호(A) = 피보험자주민번호(H)
              --   증번(A.POLI_NO) → 고객역할(G, 피보험자 '21') → 고객(H) 주민번호
              , CASE WHEN EXISTS (SELECT 1
                                    FROM EAGLE.GOGEKYUKHAL G
                                       , EAGLE.GOGEK       H
                                   WHERE G.BHGY_JGNO = A.POLI_NO   -- ※ 타입 상이 시 TO_CHAR(A.POLI_NO)
                                     AND G.HJSYYB    = '1'
                                     AND G.GGYH_CODE = '21'        -- 피보험자
                                     AND H.PK        = G.GG_PK
                                     AND H.GGNO      = A.AGNT_RSDS_NO)
                     THEN 'Y' ELSE 'N' END                    AS 본인계약여부
              -- ***유의설계사여부 : 모집사원코드(A)가 유의설계사테이블(C)에 조건으로 존재
              , CASE WHEN EXISTS (SELECT 1
                                    FROM 유의설계사테이블 C   -- ※ 실제 테이블명으로 교체
                                   WHERE C.YUI_SAWONCD = A.AGNT_CD
                                     AND C.GUBUN       = 'SU'
                                     AND C.GUBUN_S     = 'SU03')
                     THEN 'Y' ELSE 'N' END                    AS 유의설계사여부
              , A.RISK_PRM
              , A.LOSS_AMT
          FROM  DHADM.STAT_RISK_PRM_MASTER A
         WHERE  SUBSTR(A.MAIN_PROD_CODE, 1, 2) = '4A'          -- 건강보험 한정
        -- AND  A.CLOS_YR_MON = :CLOS_YR_MON                   -- 필요 시 마감년월 조건
       ) BASE
 GROUP BY BASE.계약년월
        , BASE.마감년월
        , BASE.연령대
        , BASE.담보코드
        , BASE.심사유형
        , BASE.모집지점명
        , BASE.본인계약여부
        , BASE.유의설계사여부
 ORDER BY BASE.계약년월
        , BASE.마감년월
        , BASE.연령대
        , BASE.담보코드
        , BASE.심사유형
        , BASE.모집지점명
;
