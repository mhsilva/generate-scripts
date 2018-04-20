package com.mhf.contacts.scripts;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.MessageFormat;
import java.util.List;

public class Executor {

    private static final String outputPath = "C:\\Users\\U6047034.TEN\\Desktop\\output.txt";
    private static final String excelPath = "C:\\Users\\U6047034.TEN\\Desktop\\FTA\\Additional supplier contact - layout (002).xlsx";

    public static void main(String[] args) throws Exception {
        ExcelReadingHelper excelReadingHelper = new ExcelReadingHelper();
        List<ExcelPojo> list = excelReadingHelper.createPojoFromExcel(excelPath);
        System.out.println("Total List Size " + list.size());
        StringBuilder sb = new StringBuilder();
        sb.append("BEGIN\n");
        int pos = 0;
        for (ExcelPojo excelPojo : list) {
            sb.append("DECLARE\n");
            sb.append("V_USER_ID NUMBER;\n");
            sb.append("BEGIN\n");
            sb.append("BEGIN\n");
            sb.append("INSERT INTO US_USER (TENANCY_ID, USER_ID, LOGIN_ID, USER_NM, USER_ENG_NM, DEPT_NM, POSIT_NM, POSIT_ENG_NM, EMAIL, OFFICE_TELNO, MOBLPHON_NO, FAXNO, USE_LANG_CODE, STDR_YM, EMAIL_RECPTN_YN, USE_YN, PASSWORD, PASSWORD_UPDATE_DATE, CREATE_DATE, CREATE_BY, UPDATE_DATE, UPDATE_BY, PASSWORD_FAILR_CNT, LOGIN_FAILR_YN)\n");
            sb.append(MessageFormat.format("VALUES (''OGT_FTA_MX_AD'', US_USER_S.NEXTVAL, {0}, {1}, {1}, NULL, NULL, NULL, {2}, {3}, NULL, NULL, ''en'', ''201801'', {4}, ''Y'', CRYPTIT.ENCRYPT({0}, ''FTA_'' || {0}), COMMON_PKG.GET_UTC_DATE, COMMON_PKG.GET_UTC_DATE, 18, COMMON_PKG.GET_UTC_DATE, 18, 0, ''Y'')\n", excelPojo.getLoginId(), excelPojo.getSupplierName(), excelPojo.getEmail(), excelPojo.getTelephone(), excelPojo.getReceiveEmail()));
            sb.append("RETURNING USER_ID INTO V_USER_ID;\n");
            sb.append("END;\n");
            int i = 0;
            for (String companyName : excelPojo.getCompanies()) {
                if ("".equals(companyName)) {
                    continue;
                }
                sb.append("BEGIN\n");
                //EXTR_INFO
                sb.append("BEGIN\n");
                sb.append("INSERT INTO US_USER_EXTR_INFO (TENANCY_ID, USER_ID, COMPANY_CODE, SUPPLIER_CODE, DELETE_YN, DFLT_YN, OUTPT_ORDR)\n");
                sb.append(MessageFormat.format("VALUES (''OGT_FTA_MX_AD'', V_USER_ID, (SELECT COMPANY_CODE FROM CO_COMPANY WHERE UPPER(COMPANY_NM) LIKE ''%'' || UPPER(''{0}'') || ''%'' AND TENANCY_ID = ''OGT_FTA_MX_AD''), (SELECT SUPPLIER_CODE FROM ER_SUPPLIER WHERE SUPPLIER_CODE LIKE ''%'' || {1} || ''%'' AND ROWNUM = 1), ''N'', ''N'', {2});\n", companyName, excelPojo.getSupplierCode(), i++));
                sb.append("END;\n");
                //AUTH_USER
                sb.append("BEGIN\n");
                sb.append("INSERT INTO US_AUTH_USER (TENANCY_ID, AUTH_USER_ID, USER_ID, AUTH_ID, DELETE_YN, CREATE_DATE, CREATE_BY, UPDATE_DATE, UPDATE_BY)\n");
                sb.append(MessageFormat.format("VALUES (''OGT_FTA_MX_AD'', US_AUTH_USER_S.NEXTVAL, V_USER_ID, (SELECT AUTH_ID FROM US_AUTH WHERE UPPER(COMPANY_CODE) IN (SELECT COMPANY_CODE FROM CO_COMPANY WHERE UPPER(COMPANY_NM) LIKE ''%''||UPPER(''{0}'')||''%'' AND TENANCY_ID = ''OGT_FTA_MX_AD'') AND AUTH_CODE = ''vendor'' AND TENANCY_ID = ''OGT_FTA_MX_AD''), ''Y'', COMMON_PKG.GET_UTC_DATE, 18, COMMON_PKG.GET_UTC_DATE, 18);", companyName));
                sb.append("END;\n");
                sb.append("EXCEPTION WHEN OTHERS THEN \n");
                sb.append("DBMS_OUTPUT.PUT_LINE('Deu erro no " + pos + "');\n");
                sb.append("END;\n");
            }
            sb.append("END;\n");
            sb.append("DBMS_OUTPUT.PUT_LINE('Passou pelo " + pos++ + "');\n");
        }
        sb.append("END;\n");
        Path file = Paths.get(outputPath);
        Files.write(file, sb.toString().getBytes());
    }
}
