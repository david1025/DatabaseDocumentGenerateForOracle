package tj.david.utils;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import tj.david.entity.TableInfo;

/**
 * JDBC工具类
 * @author David
 */
public class JDBCUtils {

    String driver = "oracle.jdbc.OracleDriver";
    String url = "jdbc:oracle:thin:@106.2.13.200:1521:ORCL";
    String username = "TXPFORQA";
    String password = "123456";


    public Connection getConn() {

        Connection conn = null;
        try {
            Class.forName(driver);
            conn = DriverManager.getConnection(url, username, password);

        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return conn;
    }

    public List<String> getAllTable(Connection connection) {
        List<String> list = new ArrayList<>();
        String sql = "SELECT TABLE_NAME FROM USER_TABLES";
        PreparedStatement pstmt;
        try {
            pstmt = connection.prepareStatement(sql);
            ResultSet rs = pstmt.executeQuery();
            while (rs.next()) {
                String tableName = new String(rs.getString("TABLE_NAME"));
                list.add(tableName);
            }
            pstmt.close();
            rs.close();
            connection.close();
        } catch (SQLException e) {
            e.printStackTrace();
            System.out.println("SQLException = " + e.getMessage());
        }
        return list;
    }

    public List<TableInfo> getTableColumn(Connection connection, String tableName) {

        List<TableInfo> list = new ArrayList<>();

        String sql = "SELECT " +
                " all_tab_columns.COLUMN_NAME," +
                " all_tab_columns.DATA_TYPE," +
                " all_tab_columns.DATA_LENGTH," +
                " user_col_comments.COMMENTS" +
                " FROM " +
                " all_tab_columns" +
                " LEFT JOIN " +
                " user_col_comments ON user_col_comments.COLUMN_NAME = all_tab_columns.COLUMN_NAME" +
                " where " +
                " user_col_comments.Table_Name='" + tableName + "' " +
                " and all_tab_columns.Table_Name='" + tableName + "'";
        PreparedStatement pstmt;
        try {
            pstmt = connection.prepareStatement(sql);
            ResultSet rs = pstmt.executeQuery();
            while (rs.next()) {

                TableInfo info = new TableInfo();
                info.setName(rs.getString("COLUMN_NAME"));
                info.setComment(rs.getString("COMMENTS"));
                if(info.getName().toUpperCase().equals("ID")) {
                    info.setMk("主键");
                }
                info.setType(rs.getString("DATA_TYPE"));
                list.add(info);
            }
            pstmt.close();
            rs.close();
            connection.close();
        } catch (SQLException e) {
            e.printStackTrace();
            System.out.println("SQLException = " + e.getMessage());
        }

        return list;
    }


}
