package com.export;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author: lihui
 * @date: 2018/01/31 14:59
 */
public class ExportHandle {

    public static <E> void genExcel(List<E> datas) throws IOException, IllegalAccessException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream("C:\\Users\\lenovo\\Desktop\\test.xls");
            HSSFSheet sheet = workbook.createSheet("导出excel测试");

            // 表头
            HSSFRow headRow = sheet.createRow(0);

            Map<Integer,Map<String, Field>> fieldNames = new HashMap<Integer,Map<String, Field>>();

            Class<E> clazz = (Class<E>) datas.get(0).getClass();
            Field[] fields = clazz.getDeclaredFields();
            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                ExcelExport fieldAnno = field.getAnnotation(ExcelExport.class);
                if (fieldAnno != null && fieldAnno.name() != null && !fieldAnno.name().isEmpty()) {
                    headRow.createCell(i).setCellValue(fieldAnno.name());
                    Map<String, Field> fieldMap = new HashMap<String, Field>();
                    fieldMap.put(fieldAnno.name(), field);
                    fieldNames.put(i, fieldMap);
                }
            }

            // 生成excel内容
            for (int i = 0; i < datas.size(); i++) {
                HSSFRow contentRow = sheet.createRow(i + 1);
                E e = datas.get(i);
                for (Integer index : fieldNames.keySet()) {
                    Map<String, Field> fieldMap = fieldNames.get(index);
                    Field field = fieldMap.get(fieldMap.keySet().iterator().next());
                    field.setAccessible(true);
                    Object obj = field.get(e);
                    contentRow.createCell(index).setCellValue(null2String(obj));
                    index++;
                }
            }

            workbook.write(fileOut);
        } finally {
            try {
                fileOut.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    public static String null2String(Object obj) {
        if (obj == null) {
            return "";
        }
        return obj.toString();
    }
}
