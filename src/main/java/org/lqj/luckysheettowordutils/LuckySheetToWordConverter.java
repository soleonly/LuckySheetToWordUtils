package org.lqj.luckysheettowordutils;

import com.alibaba.fastjson.JSON;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xwpf.usermodel.*;
import org.lqj.luckysheettowordutils.dto.LuckySheetDTO;
import org.lqj.luckysheettowordutils.dto.LuckySheetDataCellDTO;
import org.lqj.luckysheettowordutils.dto.LuckySheetMergeDTO;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigInteger;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class LuckySheetToWordConverter {

    public static void main(String[] args) throws FileNotFoundException {
        URL resource = LuckySheetToWordConverter.class.getClassLoader().getResource("test.docx");
        String parentPath = new File(resource.getPath()).getParent();
        String inputDocFile = parentPath+"\\test.docx";
        String luckySheetContentStr = "{\"config\":{\"merge\":{\"2_0\":{\"r\":2,\"c\":0,\"cs\":1,\"rs\":2},\"0_0\":{\"r\":0,\"c\":0,\"cs\":4,\"rs\":1},\"1_0\":{\"r\":1,\"c\":0,\"cs\":2,\"rs\":1}},\"rowlen\":{},\"columnlen\":{},\"rowhidden\":{},\"colhidden\":{},\"borderInfo\":[{\"rangeType\":\"cell\",\"value\":{\"row_index\":1,\"col_index\":2,\"l\":{\"style\":1,\"color\":\"#000\"},\"r\":{\"style\":1,\"color\":\"#000\"},\"t\":{\"style\":1,\"color\":\"#000\"},\"b\":{\"style\":1,\"color\":\"#000\"}}},{\"rangeType\":\"range\",\"borderType\":\"border-outside\",\"color\":\"#000\",\"style\":\"1\",\"range\":[{\"row\":[1,null],\"column\":[2,2]}]},{\"rangeType\":\"cell\",\"value\":{\"row_index\":1,\"col_index\":3,\"l\":{\"style\":1,\"color\":\"#000\"},\"r\":{\"style\":1,\"color\":\"#000\"},\"t\":{\"style\":1,\"color\":\"#000\"},\"b\":{\"style\":1,\"color\":\"#000\"}}},{\"rangeType\":\"range\",\"borderType\":\"border-outside\",\"color\":\"#000\",\"style\":\"1\",\"range\":[{\"row\":[1,null],\"column\":[3,3]}]},{\"rangeType\":\"range\",\"borderType\":\"border-all\",\"color\":\"#000\",\"style\":\"1\",\"range\":[{\"row\":[2,3],\"column\":[0,0]}]},{\"rangeType\":\"cell\",\"value\":{\"row_index\":2,\"col_index\":1,\"l\":{\"style\":1,\"color\":\"#000\"},\"r\":{\"style\":1,\"color\":\"#000\"},\"t\":{\"style\":1,\"color\":\"#000\"},\"b\":{\"style\":1,\"color\":\"#000\"}}},{\"rangeType\":\"cell\",\"value\":{\"row_index\":3,\"col_index\":1,\"l\":{\"style\":1,\"color\":\"#000\"},\"r\":{\"style\":1,\"color\":\"#000\"},\"t\":{\"style\":1,\"color\":\"#000\"},\"b\":{\"style\":1,\"color\":\"#000\"}}},{\"rangeType\":\"range\",\"borderType\":\"border-all\",\"color\":\"#000\",\"style\":\"1\",\"range\":[{\"row\":[1,1],\"column\":[0,1]}]},{\"rangeType\":\"cell\",\"value\":{\"row_index\":\"2\",\"col_index\":\"2\",\"l\":{\"style\":1,\"color\":\"#000\"},\"r\":{\"style\":1,\"color\":\"#000\"},\"t\":{\"style\":1,\"color\":\"#000\"},\"b\":{\"style\":1,\"color\":\"#000\"}}},{\"rangeType\":\"cell\",\"value\":{\"row_index\":\"2\",\"col_index\":\"3\",\"l\":{\"style\":1,\"color\":\"#000\"},\"r\":{\"style\":1,\"color\":\"#000\"},\"t\":{\"style\":1,\"color\":\"#000\"},\"b\":{\"style\":1,\"color\":\"#000\"}}},{\"rangeType\":\"cell\",\"value\":{\"row_index\":\"3\",\"col_index\":\"2\",\"l\":{\"style\":1,\"color\":\"#000\"},\"r\":{\"style\":1,\"color\":\"#000\"},\"t\":{\"style\":1,\"color\":\"#000\"},\"b\":{\"style\":1,\"color\":\"#000\"}}},{\"rangeType\":\"cell\",\"value\":{\"row_index\":\"3\",\"col_index\":\"3\",\"l\":{\"style\":1,\"color\":\"#000\"},\"r\":{\"style\":1,\"color\":\"#000\"},\"t\":{\"style\":1,\"color\":\"#000\"},\"b\":{\"style\":1,\"color\":\"#000\"}}}],\"authority\":{}},\"data\":[[{\"mc\":{\"r\":0,\"c\":0,\"cs\":4,\"rs\":1},\"ht\":\"0\",\"fs\":12,\"bl\":1,\"sjbs\":0,\"m\":\"指标名称、资产总计-全部-总量\",\"v\":\"指标名称、资产总计-全部-总量\",\"ct\":{\"fa\":\"General\",\"t\":\"g\"}},{\"mc\":{\"r\":0,\"c\":0},\"ht\":\"0\",\"fs\":12,\"bl\":1,\"sjbs\":0},{\"mc\":{\"r\":0,\"c\":0},\"ht\":\"0\",\"fs\":12,\"bl\":1,\"sjbs\":0},{\"mc\":{\"r\":0,\"c\":0},\"ht\":\"0\",\"fs\":12,\"bl\":1,\"sjbs\":0},null,null,null,null,null,null,null,null,null,null,null,null,null,null],[{\"mc\":{\"r\":1,\"c\":0,\"cs\":2,\"rs\":1},\"ht\":\"0\",\"fs\":10,\"bl\":0,\"sjbs\":0,\"m\":\"项目\",\"v\":\"项目\",\"ct\":{\"fa\":\"General\",\"t\":\"g\"}},{\"mc\":{\"r\":1,\"c\":0},\"ht\":\"0\",\"fs\":10,\"bl\":0,\"sjbs\":0},{\"m\":\" 指标名称\",\"v\":\" 指标名称\",\"ht\":\"0\",\"fs\":10,\"bl\":0,\"ct\":{\"fa\":\"@\",\"t\":\"s\"},\"ps\":null,\"sjbs\":0},{\"m\":\" 资产总计-全部-总量\",\"v\":\" 资产总计-全部-总量\",\"ht\":\"0\",\"fs\":10,\"bl\":0,\"ct\":{\"fa\":\"@\",\"t\":\"s\"},\"ps\":null,\"sjbs\":0},null,null,null,null,null,null,null,null,null,null,null,null,null,null],[{\"mc\":{\"r\":2,\"c\":0,\"cs\":1,\"rs\":2},\"ht\":1,\"fs\":10,\"bl\":0,\"sjbs\":0,\"m\":\"2023年\",\"v\":\"2023年\",\"ct\":{\"fa\":\"General\",\"t\":\"g\"}},{\"m\":\" 北  京\",\"v\":\" 北  京\",\"ht\":1,\"fs\":10,\"bl\":0,\"ct\":{\"fa\":\"@\",\"t\":\"s\"},\"ps\":null,\"sjbs\":0},{\"m\":\" -\",\"v\":\" -\",\"ht\":\"0\",\"fs\":10,\"bl\":0,\"ct\":{\"fa\":\"@\",\"t\":\"s\"},\"ps\":null,\"sjbs\":1},{\"m\":\" -\",\"v\":\" -\",\"ht\":\"0\",\"fs\":10,\"bl\":0,\"ct\":{\"fa\":\"@\",\"t\":\"s\"},\"ps\":null,\"sjbs\":1},null,null,null,null,null,null,null,null,null,null,null,null,null,null],[{\"mc\":{\"r\":2,\"c\":0},\"ht\":1,\"fs\":10,\"bl\":0,\"sjbs\":0},{\"m\":\" 天  津\",\"v\":\" 天  津\",\"ht\":1,\"fs\":10,\"bl\":0,\"ct\":{\"fa\":\"@\",\"t\":\"s\"},\"ps\":null,\"sjbs\":0},{\"m\":\" -\",\"v\":\" -\",\"ht\":\"0\",\"fs\":10,\"bl\":0,\"ct\":{\"fa\":\"@\",\"t\":\"s\"},\"ps\":null,\"sjbs\":1},{\"m\":\" -\",\"v\":\" -\",\"ht\":\"0\",\"fs\":10,\"bl\":0,\"ct\":{\"fa\":\"@\",\"t\":\"s\"},\"ps\":null,\"sjbs\":1},null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null],[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null]]}";
        String outputDocFile = parentPath+"\\result.docx";
        appendContentToWord(luckySheetContentStr,new FileInputStream(new File(inputDocFile)),outputDocFile);
    }

    private static String NUMBERFONTFAMILY = "Times New Roman";
    private static String STRINGFONTFAMILY = "SimSun";
    private static Integer DEFAULTFONTSIZE = 9;

    private static int getDocPageWidth(XWPFDocument doc){
        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        if (sectPr != null) {
            CTPageSz pageSize = sectPr.getPgSz();
            if (pageSize != null) {
                // 页面宽度以 TWIPS 为单位 (1 TWIP = 1/1440 inch)
                BigInteger pageWidthBigInt = (BigInteger)pageSize.getW();
                int pageWidth = pageWidthBigInt.intValue();
                return pageWidth;
            }
        }
        return 1000;
    }
    public static InputStream appendContentToWord(String luckySheetContentStr, InputStream wordFis,String outputDocFile){
        LuckySheetDTO luckySheetDTO = JSON.parseObject(luckySheetContentStr, LuckySheetDTO.class);
        try (InputStream docFis = wordFis;
             XWPFDocument doc = new XWPFDocument(docFis)) {
            // 添加一个空行
            doc.createParagraph();

            //查找sheet1中的单元格连续区域
            List<int[]> dataRanges = findContinuousRegions(luckySheetDTO.getData());
            if(CollectionUtils.isNotEmpty(dataRanges)){
                int dataRangesSize = dataRanges.size();
                for (int i = 0; i < dataRangesSize; i++) {
                    int[] range = dataRanges.get(i);
                    try {
                        // 3. 复制数据和样式到Word文档中
                        copyLuckySheetDataAndStylesToWord(doc, luckySheetDTO, range);
                    } catch (Exception e) {
                        throw new RuntimeException(e);
                    }
                }
                // 添加一个空行
                doc.createParagraph();
            }
            // 4. 保存Word文档
            saveWordDocument(doc, outputDocFile);
            return null;
            /*ByteArrayOutputStream out = new ByteArrayOutputStream();
            doc.write(out);
            // 将ByteArrayOutputStream转换为InputStream
            return new ByteArrayInputStream(out.toByteArray());*/
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static List<Integer> findWidthZeroColIndexs(XSSFSheet sheet) {
        XSSFRow excelRow = sheet.getRow(0);
        int numCols = excelRow.getPhysicalNumberOfCells();
        List<Integer> zeroColIndexs = new ArrayList<>();
        for (int colIndex = 0; colIndex < numCols; colIndex++) {
            if(sheet.getColumnWidth(colIndex) == 0){
                zeroColIndexs.add(colIndex);
            }
        }
        return zeroColIndexs;
    }
    private static void shiftCellsLeft(Row row, int columnIndexToDelete) {
        int lastColumn = row.getLastCellNum();
        for (int i = columnIndexToDelete; i < lastColumn - 1; i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            Cell nextCell = row.getCell(i + 1);
            if (nextCell != null) {
                cell.setCellValue(nextCell.getStringCellValue());
            }
        }
        if(lastColumn<1){
            return;
        }
        // 清空最后一个单元格，以便不会复制最后一个单元格的内容到上一个单元格
        Cell lastCell = row.getCell(lastColumn - 1);
        if (lastCell != null) {
            row.removeCell(lastCell);
        }
    }

    private static int nullCellStartIndex(List<LuckySheetDataCellDTO> cells){
        int continuousNullCellCount = 0;
        if(CollectionUtils.isEmpty(cells)){
            return 0;
        }
        int result = 0;
        int size = cells.size();
        for (int i = 0; i < size; i++) {
            if(continuousNullCellCount >3){
                break;
            }
            if(cells.get(i) == null){
                if(continuousNullCellCount == 0){
                    result = i;
                }
                continuousNullCellCount++;
            }
        }
        return result;
    }

    public static List<int[]> findContinuousRegions(List<List<LuckySheetDataCellDTO>> data) {
        List<int[]> regions = new ArrayList<>();
        int totalRowCount = data.size();
        int endRow = 0,endCol = 0;
        for (int r = 0; r < totalRowCount; r++) {
            List<LuckySheetDataCellDTO> row = data.get(r);
            int nullCellStartIndex = nullCellStartIndex(row);
            if (row != null && nullCellStartIndex >0) {
                if(endCol == 0){
                    endCol = nullCellStartIndex;
                }
                endRow = r;
            }else{
                break;
            }
        }
        regions.add(new int[]{0, 0, endRow, endCol-1});
        return regions;
    }

    /**
     * 从Excel复制数据和样式到Word文档
     */
    private static void copyLuckySheetDataAndStylesToWord(XWPFDocument doc, LuckySheetDTO luckySheet, int[] dataRange) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        int startRow = Integer.valueOf(dataRange[0]);
        int endRow = Integer.valueOf(dataRange[2]);
        int startCol = Integer.valueOf(dataRange[1]);
        int endCol = Integer.valueOf(dataRange[3]);
        int docPageWidth = getDocPageWidth(doc);
        XWPFTable table = doc.createTable(endRow-startRow+1,endCol-startCol+1); // 创建新的表格
//        table.setInsideHBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
//        table.setInsideVBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
//        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
//        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
//        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
//        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
        List<List<LuckySheetDataCellDTO>> sheetData = luckySheet.getData();
        Set<String> cellMergedRegions = new HashSet<>();
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            XWPFTableRow wordRow = table.getRow(rowIndex-startRow); // 创建新的表格行
            if (wordRow == null) {
                wordRow = table.createRow();
            }

            List<LuckySheetDataCellDTO> excelRow = sheetData.get(rowIndex);

            //wordRow.setHeight((int) (excelRow.getHeight())); // 将 Excel 高度转换为 Twips 单位

            //整理行中列宽度数组
            int numCols = endCol+1;
            int[] colWidths = null;
            if(colWidths == null){
                colWidths = new int[numCols];
                int colWidth = docPageWidth/numCols;
                for (int colIndex = 0; colIndex < numCols; colIndex++) {
                    colWidths[colIndex] = colWidth;
                }
            }

            int deleteColIndexCount = 0;
            for (int colIndex = startCol; colIndex <= endCol; colIndex++) {
                int wordCellIndex = colIndex-startCol-deleteColIndexCount;
                int colWidth = colWidths[colIndex];
                XWPFTableCell wordCell = wordRow.getCell(wordCellIndex); // 创建新的表格单元格
                if (wordCell == null) {
                    wordCell = wordRow.createCell();
                }
                LuckySheetDataCellDTO luckyCell = excelRow.get(colIndex);
                if (luckyCell != null) {
                    // 获取单元格的内容和样式
                    copyCellValueAndStyle(luckyCell, colWidth, wordCell, rowIndex == 0);
                }
            }
        }
        if (luckySheet.getConfig() != null && luckySheet.getConfig().getMerge() != null && (!luckySheet.getConfig().getMerge().isEmpty())) {
            luckySheet.getConfig().getMerge().values().forEach(m->{
                String cellMergedRegion =  generateCellMargin(m);
                cellMergedRegions.add(cellMergedRegion);
            });
            if(CollectionUtils.isNotEmpty(cellMergedRegions)){
                cellMergedRegions.forEach(cms->{
                    String[] splits = cms.split(",");
                    mergeCells(table,Integer.valueOf(splits[0]),Integer.valueOf(splits[1]),Integer.valueOf(splits[2]),Integer.valueOf(splits[3]));
                });
            }
        }
    }

    private static String generateCellMargin(LuckySheetMergeDTO mergeDTO) {
        int rowStart = mergeDTO.getR();
        int rowEnd = mergeDTO.getR() + mergeDTO.getRs()-1;
        int colStart = mergeDTO.getC();
        int colEnd = mergeDTO.getC() + mergeDTO.getCs()-1;
        return String.format("%d,%d,%d,%d", rowStart, colStart, rowEnd, colEnd);
    }

    private static int generateMinusCount(int colIndex, List<Integer> zeroColIndexs) {
        if(CollectionUtils.isEmpty(zeroColIndexs)){
            return 0;
        }
        int count = 0;
        for (Integer zeroColIndex : zeroColIndexs) {
            if(zeroColIndex<colIndex){
                count++;
            }
        }
        return count;
    }

    private static boolean isCellInHiddenMergedRegion(Sheet sheet, Cell cell) {
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();

        for (CellRangeAddress region : mergedRegions) {
            if (region.isInRange(rowIndex, columnIndex)) {
                // Check if the merged region is hidden (either row or column is hidden)
                for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
                    if (sheet.getRow(i).getZeroHeight()) {
                        return true;
                    }
                }
                for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                    if (sheet.isColumnHidden(j)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    public static void mergeCells(XWPFTable table, int top, int left, int bottom,int right) {
        for(int i= top;i<=bottom;i++){
            mergeCellsHorizontal(table,i,left,right);
        }
        for(int i= left;i<=right;i++){
            mergeCellsVertically(table,i,top,bottom);
        }
    }


    // 水平合并单元格
    public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        XWPFTableRow tableRow = table.getRow(row);
        if(tableRow == null){
            return;
        }
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = tableRow.getCell(cellIndex);
            if(cell==null){
                continue;
            }
            if (cellIndex == fromCell) {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if(row == null){
                continue;
            }
            XWPFTableCell cell = row.getCell(col);
            if(cell==null){
                continue;
            }
            if (rowIndex == fromRow) {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    // 获取单元格合并区域
    private static CellRangeAddress getMergedRegion(Sheet sheet, int rowIdx, int colIdx) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowIdx, colIdx)) {
                return merged;
            }
        }
        return null;
    }

    /**
     * 获取单元格的内容
     */
    private static String getCellValue(LuckySheetDataCellDTO cell) {
        return cell.getV();
    }


    private static void setCellWidthForHidden(XWPFTableCell wordCell){
        CTTc ctTc = wordCell.getCTTc();
        CTTcPr tcPr = ctTc.addNewTcPr();
        CTTblWidth tblWidth = tcPr.addNewTcW();
        tblWidth.setW(BigInteger.valueOf(0));

        // 设置单元格边框颜色为白色
        CTTcBorders borders = tcPr.addNewTcBorders();
        borders.addNewTop().setColor("FFFFFF");
        borders.addNewBottom().setColor("FFFFFF");
        borders.addNewLeft().setColor("FFFFFF");
        borders.addNewRight().setColor("FFFFFF");

        // 添加一个空段落并设置其颜色为白色
        for (XWPFParagraph paragraph : wordCell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                run.setColor("FFFFFF");
            }
        }
    }
    //设置单元格宽度
    private static void setCellWidth(XWPFTableCell wordCell,int colWidth){
        CTTcPr tcPr = wordCell.getCTTc().getTcPr();
        if (tcPr == null) {
            tcPr = wordCell.getCTTc().addNewTcPr();
        }
        CTTblWidth tblWidth = tcPr.getTcW();
        if (tblWidth == null) {
            tblWidth = tcPr.addNewTcW();
        }
        tblWidth.setW(BigInteger.valueOf(colWidth)); // 设置单元格宽度为0
    }
    private static Pattern numberPattern = Pattern.compile("^-?\\d+(\\.\\d+)?$");
    /**
     * 复制单元格样式
     */
    private static void copyCellValueAndStyle(LuckySheetDataCellDTO luckyCell,int colWidth, XWPFTableCell wordCell,boolean firstRow) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        String cellValue = getCellValue(luckyCell);
        if(cellValue == null){
            return;
        }
        XWPFParagraph paragraph = wordCell.getParagraphs().get(0);
        XWPFRun run = paragraph.createRun();
        cellValue = cellValue.replace("\n", "").replace("\r"," ").replace("\t","").trim();
        // 复制内容到Word文档中
        run.setText(cellValue);
        Matcher matcher = numberPattern.matcher(cellValue);
        boolean isNumber = matcher.matches();

        //设置单元格宽度
        setCellWidth(wordCell,colWidth);

        // 通过反射调用受保护的 getCellAlignment 方法
        Method getCellAlignmentMethod = XSSFCellStyle.class.getDeclaredMethod("getCellAlignment");
        getCellAlignmentMethod.setAccessible(true);
        // 获取水平对齐和垂直对齐方式
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        CTTcPr ctTcPr = getCTTcPr(wordCell);
        ctTcPr.addNewVAlign().setVal(STVerticalJc.CENTER);

        // 设置字体样式
        CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        CTFonts fonts = rpr.addNewRFonts();
        if(isNumber){
            run.setFontFamily(NUMBERFONTFAMILY);
            fonts.setEastAsia(NUMBERFONTFAMILY);
            fonts.setAscii(NUMBERFONTFAMILY);
            fonts.setHAnsi(NUMBERFONTFAMILY);
        }else{
            run.setFontFamily(STRINGFONTFAMILY);
            fonts.setEastAsia(STRINGFONTFAMILY);
            fonts.setAscii(STRINGFONTFAMILY);
            fonts.setHAnsi(STRINGFONTFAMILY);
        }
        run.setFontSize(DEFAULTFONTSIZE);
        if(firstRow){
            run.setBold(true);
        }

        // 设置单元格背景色
        /*if (sourceStyle.getFillPattern() == FillPatternType.SOLID_FOREGROUND) {
            wordCell.setColor(getXWPFColor(sourceStyle.getFillForegroundColorColor()));
        }*/

        // 设置边框
        setCellBorders(luckyCell, wordCell, firstRow);
    }

    /**
     * 获取XWPF颜色对象
     */
    private static String getXWPFColor(Color color) {
        if (color instanceof XSSFColor) {
            byte[] rgb = ((XSSFColor) color).getRGB();
            if (rgb != null) {
                int red = rgb[0] & 0xFF;
                int green = rgb[1] & 0xFF;
                int blue = rgb[2] & 0xFF;
                return String.format("%02X%02X%02X", red, green, blue);
            }
        }
        return null;
    }

    private static CTTcPr getCTTcPr(XWPFTableCell targetCell){
        CTTc ctTc = targetCell.getCTTc();
        CTTcPr ctTcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
        return ctTcPr;
    }
    private static void removeTableBorders(XWPFTable table) {
        // 设置表格的整体边框为无
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr == null) {
            tblPr = table.getCTTbl().addNewTblPr();
        }
        CTTblBorders tblBorders = tblPr.getTblBorders();
        if (tblBorders == null) {
            tblBorders = tblPr.addNewTblBorders();
        }
        tblBorders.addNewTop().setVal(STBorder.NONE);
        tblBorders.addNewBottom().setVal(STBorder.NONE);
        tblBorders.addNewLeft().setVal(STBorder.NONE);
        tblBorders.addNewRight().setVal(STBorder.NONE);
        tblBorders.addNewInsideH().setVal(STBorder.NONE);
        tblBorders.addNewInsideV().setVal(STBorder.NONE);

        // 设置每个单元格的边框为无
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                CTTcPr tcPr = cell.getCTTc().getTcPr();
                if (tcPr == null) {
                    tcPr = cell.getCTTc().addNewTcPr();
                }
                CTTcBorders borders = tcPr.isSetTcBorders() ? tcPr.getTcBorders() : tcPr.addNewTcBorders();
                borders.addNewTop().setVal(STBorder.NONE);
                borders.addNewBottom().setVal(STBorder.NONE);
                borders.addNewLeft().setVal(STBorder.NONE);
                borders.addNewRight().setVal(STBorder.NONE);
                borders.addNewInsideH().setVal(STBorder.NONE);
                borders.addNewInsideV().setVal(STBorder.NONE);
            }
        }
    }
    /**
     * 设置单元格边框
     */
    private static void setCellBorders(LuckySheetDataCellDTO luckyCell, XWPFTableCell cell,boolean firstRow) {
        CTTc ctTc = cell.getCTTc();
        if(firstRow){
            CTTcBorders ctTcBorders = ctTc.addNewTcPr().addNewTcBorders();
            ctTcBorders.addNewTop().setVal(STBorder.NONE);
            ctTcBorders.addNewBottom().setVal(STBorder.NONE);
            ctTcBorders.addNewLeft().setVal(STBorder.NONE);
            ctTcBorders.addNewRight().setVal(STBorder.NONE);
        }else{
            CTTcBorders ctTcBorders = ctTc.addNewTcPr().addNewTcBorders();
            ctTcBorders.addNewTop().setVal(STBorder.SINGLE);
            ctTcBorders.addNewBottom().setVal(STBorder.SINGLE);
            ctTcBorders.addNewLeft().setVal(STBorder.SINGLE);
            ctTcBorders.addNewRight().setVal(STBorder.SINGLE);
        }

    }


    /**
     * 保存Word文档
     */
    /*private static void outputWordDocument(HttpServletResponse response, XWPFDocument doc) throws IOException {

        try {
            response.addHeader("Content-Disposition", "attachment;filename=" + "test.doc");
//            response.setHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes("GB2312"), "utf-8") + System.currentTimeMillis() + ".xls");
            //设置文本格式
            response.setCharacterEncoding("utf-8");
            response.setContentType("application/msexcel");
            doc.write(response.getOutputStream());
            response.getOutputStream().flush();
            response.getOutputStream().close();
        } finally {
            doc.close();
        }
    }*/

    /**
     * 保存Word文档
     */
    private static void saveWordDocument(XWPFDocument doc, String outputFile) throws IOException {
        FileOutputStream out = null;
        try {
            File file = new File(outputFile);
            if(!file.exists()){
                file.createNewFile();
            }
            out = new FileOutputStream(file);
            doc.write(out);
        } finally {
            if (out != null) {
                out.close();
            }
            doc.close();
        }
    }
}
