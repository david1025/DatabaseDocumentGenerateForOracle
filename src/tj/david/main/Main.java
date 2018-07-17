package tj.david.main;

import java.io.FileOutputStream;
import java.sql.Connection;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import tj.david.entity.TableInfo;
import tj.david.utils.HSSF2BeanColumMatch;
import tj.david.utils.JDBCUtils;


/**
 * 程序运行主类
 * @author David
 */
public class Main {

    public static void main(String[] args) {

        XSSFWorkbook workbook;
        try {
            //创建一个XSSFWorkbook
            workbook = new XSSFWorkbook();

            List<String> tables = getTables();

            for (String string : tables) {
                //根据表明创建sheet
                workbook.createSheet(string);

                List<TableInfo> list = new ArrayList<TableInfo>();

                //new一个TableInfo当作表头，以后扩展自定义标题和边框线
                TableInfo tableInfo = new TableInfo();
                tableInfo.setName("字段");
                tableInfo.setComment("说明");
                tableInfo.setType("数据类型");
                tableInfo.setMk("类型");
                list.add(tableInfo);

                List<TableInfo> list1 = getTableInfos(string);

                list.addAll(list1);

                //创意一个HSSF2BeanColumMatch
                HSSF2BeanColumMatch<TableInfo> h2bc = new HSSF2BeanColumMatch<TableInfo>(TableInfo.class);
                //此处MatchFieldName需要和你的实体的属性相对应
                h2bc.setMatchFieldNames(new String[]{"name", "type", "comment", "mk"});
                //从第0行开始写数据，到list.size()行
                h2bc.setStatEndNum(0, list.size());
                h2bc.bean2xlsx(workbook, string, list);
                //如果有公式，需要调用此方法刷新一下
                workbook.getSheet(string).setForceFormulaRecalculation(true);
            }

            FileOutputStream fos;

            //替换成你想要输出的文件名
            fos = new FileOutputStream("D:/aa.xlsx");
            workbook.write(fos);
            fos.close();

        } catch (Exception e) {
            e.printStackTrace();

        }
    }

    private static List<String> getTables() {
        JDBCUtils jdbcUtils = new JDBCUtils();
        Connection con = jdbcUtils.getConn();
        return jdbcUtils.getAllTable(con);
    }


    private static List<TableInfo> getTableInfos(String tableName) {

        JDBCUtils jdbcUtils = new JDBCUtils();
        Connection con = jdbcUtils.getConn();

        return jdbcUtils.getTableColumn(con, tableName);
    }

}
