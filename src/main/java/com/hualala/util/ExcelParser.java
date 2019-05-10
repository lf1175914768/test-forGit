package com.hualala.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * @Description TODO
 * @Author liufeng
 * @Date 2019/5/7  9:37
 */
public class ExcelParser {

    private final String DEFAULT_CONFIG_PATH = "config.properties";
    private final String DEFAULT_MAPPING_PATH = "mapping.properties";
    private final String DEFAULT_CONVERT_MAPPING_PATH = "convertMapping.properties";
    private final String DEFAULT_SEPERATOR = ",";
    private String separator = DEFAULT_SEPERATOR;

    private Properties prop = new Properties();
    private Set<String> ignoreFields = new HashSet<>();

    public ExcelParser() {
        load(DEFAULT_CONFIG_PATH);
    }

    private void load(String path) {
        load0(path, prop);
        String field = prop.getProperty("ignoreFields");
        if(field != null) {
            String[] fields = field.split(",");
            for(String can : fields) {
                if(can.trim().length() > 0) {
                    ignoreFields.add(can);
                }
            }
        }
    }

    private void load0(String path, Properties properties) {
        InputStream stream = ExcelParser.class.getClassLoader().getResourceAsStream(path);
        if(stream == null) {
            throw new IllegalArgumentException("The path is not correct, please check the path.");
        }
        Reader reader = null;
        try {
            reader = new InputStreamReader(stream, "UTF-8");
            properties.load(reader);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if(reader != null) {
                try {
                    reader.close();
                } catch (IOException e) {
                }
            }
        }
    }

    public void parse(String path) {
        if(path == null) {
            path = prop.getProperty("path");
        }
        File file = new File(path);
        if(!file.exists()) {
            throw new IllegalArgumentException("The file path you passed is not found. ...");
        }
        String [] pieces = file.getName().split("\\.");
        Workbook book = null;
        try {
            if("xls".equals(pieces[1])) {
                FileInputStream fis = new FileInputStream(file);
                book = new HSSFWorkbook(fis);
            } else if("xlsx".equals(pieces[1])) {
                book = new XSSFWorkbook(file);
            } else {
                System.out.println("The file Type error");
                return;
            }
            doParseInternal(book);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } finally {
            if(book != null) {
                try {
                    book.close();
                } catch (IOException e) {
                }
            }
        }
    }

    private void doParseInternal(Workbook book) {
        Sheet sheet = book.getSheetAt(0);   //Default parse The first sheet
        Internal internal = new Internal(prop.getProperty("db"),prop.getProperty("table"));
        internal.prefixOfField = prop.getProperty("prefixOfField") == null ? "" : prop.getProperty("prefixOfField");
        internal.suffixOfField = prop.getProperty("suffixOfField") == null ? "" : prop.getProperty("suffixOfField");

        internal.setFirstRow(prop.getProperty("headerRow"));
        internal.setActualFirstRow(prop.getProperty("headerRowOffset"));
        Map<String, Integer> mappings = generateFieldMapping(sheet.getRow(internal.firstRow));

        internal.prefixGenerated = prop.getProperty("generatedFieldPrefix") == null ? "" : prop.getProperty("generatedFieldPrefix");
        internal.suffixGenerated = prop.getProperty("generatedFieldSuffix") == null ? "" : prop.getProperty("generatedFieldSuffix");
        internal.setGeneratedFields(prop.getProperty("generatedFields"), prop.getProperty("generatedStartIndex"));
        internal.convertMapping = parseConvertMapping(DEFAULT_CONVERT_MAPPING_PATH);

        StringBuilder sb = new StringBuilder();
        StringBuilder temp = new StringBuilder();
        String fields = generateFields(mappings.keySet(), internal, temp);
        for(int i = internal.actualFirstRow; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            sb.append(prop.get("type") + " ");
            sb.append(internal.getTable() + " ");
            sb.append(internal.prefixOfField + fields + internal.suffixOfField);
            sb.append(" VALUES ");
            sb.append(internal.prefixOfField + generateValues(mappings, row, temp, internal) + internal.suffixOfField);
            sb.append(";\n");
//            System.out.println(row.getCell(0).getStringCellValue());
        }
        System.out.println(mappings);
        System.out.println(sb.toString());

    }

    private String generateValues(Map<String, Integer> mappings, Row row, StringBuilder sb, Internal internal) {
        sb.delete(0, sb.length());
        for(Map.Entry<String, Integer> entry : mappings.entrySet()) {
            sb.append("'" + row.getCell(entry.getValue()) + "'").append(", ");
        }
        for(Map.Entry<String, Map<String, Object>> entry : internal.convertMapping.entrySet()) {
            Map<String, Object> entryMap = entry.getValue();
            Object obj = entryMap.get(row.getCell(mappings.get(decompileIdentifier(entry.getKey(), 0))).toString());
            sb.append("'" + obj + "'").append(", ");
        }
        if(internal.countGeneratedField > 0) {
            // Add Generated Values
            for(Map.Entry<String ,Integer> entry : internal.generatedFields.entrySet()) {
                sb.append("'" + internal.prefixGenerated + entry.getValue() + internal.suffixGenerated + "'").append(", ");
                entry.setValue(entry.getValue() + 1);
            }
        }
        return sb.delete(sb.length() - 2, sb.length()).toString();
    }

    private Map<String,Integer> generateFieldMapping(Row row) {
        Properties mapping = new Properties();
        load0(DEFAULT_MAPPING_PATH, mapping);
        Map<String, Integer> result = new HashMap<>();
        for(int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            String key;
            if(cell == null || (key = cell.getStringCellValue()).length() == 0)
                throw new IllegalStateException("The Header At the " + i + " index is empty, Please remove or add something to the header");
            String value = mapping.getProperty(key);
            if(value != null) {
                putValues(value, result, i);
            } else if(!ignoreFields.contains(key)) {
                result.put(key, i);
            }
        }
        return result;
    }

    private void putValues(String value, Map<String, Integer> result, int index) {
        if(value.length() == 0) {
            throw new IllegalArgumentException("The DB Field is not allowed to be empty...");
        }
        String[] pieces = value.split(",");
        for(String piece : pieces) {
            if(piece.trim().length() > 0) {
                result.put(piece, index);
            }
        }
    }

    private String generateFields(Set<String> sets, Internal obj, StringBuilder temp) {
        Map<String, Integer> generatedFields = obj.generatedFields;
        Set<String> convertMappingSet = obj.convertMapping.keySet();
        temp.delete(0, temp.length());
        for(String candidate : sets) {
            temp.append(candidate).append(", ");
        }
        for(String candidate : convertMappingSet) {
            temp.append(decompileIdentifier(candidate, 1)).append(", ");
        }
        if(generatedFields != null) {
            Set<String> pieces = generatedFields.keySet();
            for(String piece : pieces) {
                if(piece.trim().length() > 0) {
                    temp.append(piece).append(", ");
                }
            }
        }
        return temp.delete(temp.length() - 2, temp.length()).toString();
    }

    // 还没想好用什么 特殊字符来表示关闭 Start，暂时只要 有一个，就一直开着
    private Map<String, Map<String, Object>> parseConvertMapping(String path) {
        Map<String, Map<String, Object>> result = new HashMap<>();
        InputStream inStream = ExcelParser.class.getClassLoader().getResourceAsStream(path);
        if(inStream == null) {
            throw new IllegalArgumentException("The path is not correct, please check the path.");
        }
        BufferedReader reader = null;
        try {
            reader = new BufferedReader(new InputStreamReader(inStream, "UTF-8"));
            String line, currentIdentifier = null;
            boolean start = false;
            while((line = reader.readLine()) != null) {
                line = line.trim();
                if(line.startsWith("#") || line.length() <= 0)
                    ///  it is comments or empty, skip it.
                    continue;
                if(line.contains("=")) {
                    Map<String, Object> candidate = new HashMap<>();
                    result.put((currentIdentifier = generateIdentifier(line)), candidate);
                    start = true;
                    continue;
                }
                if(start) {
                    if(currentIdentifier == null || currentIdentifier.length() <= 0) {
                        throw new IllegalStateException("The switch is on, but we didn't know what currentIdentifier is...");
                    }
                    Map<String, Object> rs = result.get(currentIdentifier);
                    String[] ss = line.split(separator);
                    if(ss.length != 2) {
                        throw new IllegalStateException("The ConvertMapping must have two element , with separator '[ " + separator + " ];");
                    }
                    rs.put(canonicalize(ss[0]), canonicalize(ss[1]));
                }
            }
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if(reader != null) {
                    reader.close();
                }
            } catch (IOException e) {
            }
        }
        return result;
    }

    private String canonicalize(String elem) {
        String rs = elem.trim(); int pos;
        if(elem.startsWith("\""))
            rs = rs.substring(1, (pos = rs.lastIndexOf("\"")) < 1 ? rs.length() : pos);
        else if(elem.startsWith("'"))
            rs = rs.substring(1, (pos = rs.lastIndexOf("'")) < 1 ? rs.length() : pos);
        return rs;
    }

    /// 中间用$$$ 分隔
    private String generateIdentifier(String line) {
        int index = line.indexOf("=");
        if(index < 0) {
            throw new IllegalArgumentException("The parameter line you passed '" + line + "' doesn't have '=', Error....");
        }
        return line.substring(0, index) + "$$$" + line.substring(index + 1);
    }

    private String decompileIdentifier(String identifier, int i) {
        String[] strs = identifier.split("\\$\\$\\$");
        if(strs.length < 2) {
            throw new IllegalArgumentException("The parameter line you passed '" + identifier + "' doesn't have '$$$', Error....");
        }
        return strs[i];
    }

    private class Internal {
        String prefixGenerated;
        String suffixGenerated;
        String db;
        String table;
        String prefixOfField;
        String suffixOfField;
        int firstRow;
        int actualFirstRow;
        Map<String, Integer> generatedFields;
        int countGeneratedField;
        Map<String, Map<String, Object>> convertMapping;

        Internal(String db, String table) {
            this.db = db;
            if(table == null || table.length() <= 0) {
                throw new IllegalStateException("The table must configure, Please check and retry...");
            }
            this.table = table;
        }

        String getTable() {
            if(db != null && db.length() > 0) {
                return db + "." + table;
            }
            return table;
        }
        void setFirstRow(String firstRow) {
            if(firstRow != null && firstRow.trim().length() > 0) {
                this.firstRow = Integer.parseInt(firstRow);
            }
        }
        void setActualFirstRow(String firstRowOffset) {
            if(firstRowOffset != null && firstRowOffset.trim().length() > 0) {
                int firstRowOff = Integer.parseInt(firstRowOffset);
                if(firstRowOff > 1) {
                    throw new IllegalStateException("Not support multiple header rows for now. please check the configuration and the file...");
                }
                this.actualFirstRow = firstRow + firstRowOff;
            } else {
                this.actualFirstRow = firstRow + 1;
            }
        }
        void setGeneratedFields(String generatedFields, String generatedIndexs) {
            if(generatedFields != null && generatedFields.length() > 0) {
                if(generatedIndexs == null || generatedIndexs.length() <= 0) {
                    throw new IllegalStateException("The generatedStartIndex is not set, in case of generatedFields is set, Please check and retry...");
                }
                String[] fields = generatedFields.split(",");
                String[] indexes = generatedIndexs.split(",");
                if(indexes == null || indexes.length <= 0 || indexes[0].trim().length() <= 0) {
                    throw new IllegalStateException("The First generatedStartIndex must set, Please configure it and retry...");
                }
                this.generatedFields = new HashMap<>(fields.length);
                int prev = 0, minLen = Math.min(fields.length, indexes.length);

                for(int i = 0; i < minLen; i++) {
                    if(indexes[i].trim().length() <= 0) {
                        this.generatedFields.put(fields[i].trim(), prev);
                        continue;
                    }
                    this.generatedFields.put(fields[i].trim(), (prev = Integer.parseInt(indexes[i].trim())));
                }

                if(minLen < fields.length) {
                    for(int i = minLen; i < fields.length; i++) {
                        this.generatedFields.put(fields[i].trim(), prev);
                    }
                }
                this.countGeneratedField = fields.length;
            }
        }
    }
}
