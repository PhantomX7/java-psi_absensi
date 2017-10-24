/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package absensi.psi;

import datechooser.beans.DateChooserCombo;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.Date;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.concurrent.ExecutionException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.joda.time.Days;
import org.joda.time.Period;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;
import util.DbConn;

/**
 *
 * @author Lab04
 */
public class FrmMain extends javax.swing.JFrame {

    private String path = null;//open
    private String path2 = null;//save

    private Employee employee = null;

    private Connection myConn = null;
    private PreparedStatement myStmt = null;
    private ResultSet myRs = null;

    private final ArrayList<String> nimList = new ArrayList<>();

    private ArrayList<Object[]> objectList = null;

    private LoadWorker worker;

    DateTimeFormatter formatter = DateTimeFormat.forPattern("HH:mm:ss");

    public FrmMain() {
        initComponents();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel", "xlsx");
        fchOpen.setFileFilter(filter);
        setLocationRelativeTo(null);
        connectToDatabase();

    }

    private void connectToDatabase() {
        try {
            Class.forName(DbConn.JDBC_CLASS);
            myConn = DriverManager.getConnection(DbConn.JDBC_URL,
                    DbConn.JDBC_USERNAME,
                    DbConn.JDBC_PASSWORD);

        } catch (SQLException ex) {
            showError("sql Exception");
        } catch (ClassNotFoundException ex) {
            Logger.getLogger(FrmMain.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void openAndSaveToDatabase() {
        // set label path
        lblPath.setText("xls file: " + path);
        lblPath2.setText("xls file: " + path);

        try {
            FileInputStream fileInputStream;
            fileInputStream = new FileInputStream(path);
            XSSFWorkbook workBook = new XSSFWorkbook(fileInputStream);
            for (int x = 0; x < workBook.getNumberOfSheets(); x++) {
                XSSFSheet workSheet = workBook.getSheetAt(x);
                Iterator rowIter = workSheet.rowIterator();

                int row = 0;
                ArrayList<Employee> list = new ArrayList<>();

                while (rowIter.hasNext()) {
                    XSSFRow myRow = (XSSFRow) rowIter.next();
                    
                    Iterator cellIter = myRow.cellIterator();
                    if (row == 0) {
                        row++;
                        continue;
                    }
                    //Vector cellStoreVector=new Vector();
                    employee = new Employee();
                    int i = 1;
                    while (cellIter.hasNext()) {
                        XSSFCell myCell = (XSSFCell) cellIter.next();
                        if (i == 1) {
                            employee.setDate(myCell.toString());
                            i++;
                            continue;
                        }
                        if (i == 2) {
                            employee.setTime(myCell.toString());
                            i++;
                            continue;
                        }
                        if (i == 3) {
                            employee.setId(myCell.toString());
                            employee.mergeDateAdnTime();
                        }
                    }
                    list.add(employee);

                }
                list.trimToSize();

                for (int i = 0; i < list.size(); i++) {
                    // Prepare statement
                    lblSheetMain1.setText("Sheet " + (x + 1) + "/" + workBook.getNumberOfSheets());
                    lblSheetMain2.setText("Sheet " + (x + 1) + "/" + workBook.getNumberOfSheets());
                    lblSheetMain3.setText("Sheet " + (x + 1) + "/" + workBook.getNumberOfSheets());

                    double percentage = (double) i * ((double) 100 / (list.size() - 1));
                    pgbMain.setValue((int) percentage);
                    pgbMain2.setValue((int) percentage);
                    pgbMain3.setValue((int) percentage);
                    myStmt = myConn.prepareStatement("insert into data_absensi values (?,?)");

                    SimpleDateFormat formatDate = new SimpleDateFormat("yyyyMMddHHmmss");
                    Date date = formatDate.parse(list.get(i).getDateTime());

                    java.util.Date utilDate = new java.util.Date();
                    java.sql.Timestamp timestamp = new java.sql.Timestamp(date.getTime());

                    myStmt.setTimestamp(1, timestamp);
                    myStmt.setString(2, list.get(i).getId());

                    // Execute SQL query
                    myStmt.executeUpdate();
                }
            }

        } catch (FileNotFoundException ex) {
            Logger.getLogger(FrmMain.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException | EncryptedDocumentException | SQLException ex) {
            showError("IO/sql exception");
        } catch (ParseException ex) {

        }
    }

    private void loadResultDataFromDatabase(String nik, String date, JTable table) {

        DefaultTableModel tableModel = (DefaultTableModel) table.getModel();
//        try {
//            myStmt = myConn.prepareStatement("select nik,First_Name,Last_Name,Department,date(date) as date,min(time(date))"
//                    + " as sign_in,max(time(date)) as sign_out,timediff((select max(time(date))"
//                    + " from full_data where nik=? and date(date)=?),"
//                    + "(select min(time(date)) from full_data where nik=? and date(date)=?))"
//                    + " as working_hours  from full_data where nik=? and date(date)=? ");
//
//            myStmt.setString(1, nik);
//            myStmt.setString(2, date);
//            myStmt.setString(3, nik);
//            myStmt.setString(4, date);
//            myStmt.setString(5, nik);
//            myStmt.setString(6, date);
//            // Execute statement
//            myRs = myStmt.executeQuery();
//            // Process result set
//            while (myRs.next()) {
//
//                Object data[] = {nik, myRs.getString("first_name"), myRs.getString("last_name"), myRs.getString("Department"),
//                    myRs.getString("date"), myRs.getString("sign_in"), myRs.getString("sign_out"), myRs.getString("working_hours")};
//                tableModel.addRow(data);
//            }
//        } catch (SQLException ex) {
//            showError("sql Exception");
//        }

        ArrayList<Object> data = new ArrayList<>();

        try {
            myStmt = myConn.prepareStatement("select * from data_karyawan where nik=?;");
            myStmt.setString(1, nik);
            myRs = myStmt.executeQuery();
            data.add(nik);
            if (myRs.isBeforeFirst()) {
                while (myRs.next()) {
                    System.out.println(myRs.getString("first_name"));
                    data.add(myRs.getString("first_name"));
                    data.add(myRs.getString("last_name"));
                    data.add(myRs.getString("Department"));
                }
            } else {
                data.add("");
                data.add("");
                data.add("");
            }
            data.add(date);

            myStmt = myConn.prepareStatement("select min(time(date))"
                    + "as sign_in,max(time(date)) as sign_out,timediff((select max(time(date))"
                    + "from full_data where nik=? and date(date)=?),"
                    + "(select min(time(date)) from full_data where nik=? and date(date)=?))"
                    + " as working_hours  from full_data where nik=? and date(date)=? ");
            myStmt.setString(1, nik);
            myStmt.setString(2, date);
            myStmt.setString(3, nik);
            myStmt.setString(4, date);
            myStmt.setString(5, nik);
            myStmt.setString(6, date);
            myRs = myStmt.executeQuery();

            DateTime d2 = null;
            DateTime d3 = null;
            while (myRs.next()) {
                d2 = formatter.parseDateTime(myRs.getString("sign_in"));
                d3 = formatter.parseDateTime(myRs.getString("working_hours"));
                data.add(myRs.getString("sign_in"));
                data.add(myRs.getString("sign_out"));
                data.add(myRs.getString("working_hours"));
            }
            String waktuTelat = getSignInTimeLimit(nik);

            if (waktuTelat != null) {
                DateTime d1 = formatter.parseDateTime(waktuTelat);
                Period period = new Period(d1, d2);
                if (period.getHours() > 0) {
                    data.add(true);
                } else {
                    if (period.getMinutes() > 5) {
                        data.add(true);
                    } else {
                        data.add(false);
                    }
                }
            } else {
                data.add(false);
            }

            if (d3.getHourOfDay() >= 9) {
                data.add(true);
            } else {
                data.add(false);
            }

            tableModel.addRow(data.toArray());
        } catch (SQLException ex) {
            Logger.getLogger(FrmMain.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private String getSignInTimeLimit(String nik) {
        try {
            myStmt = myConn.prepareStatement("select jam_masuk from data_karyawan where nik=?;");
            myStmt.setString(1, nik);
            myRs = myStmt.executeQuery();
            while (myRs.next()) {
                return myRs.getString("jam_masuk");
            }

        } catch (SQLException ex) {
            Logger.getLogger(FrmMain.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }

    //dosen tidak tetap
    private void loadResultDataFromDatabase2(String nik, String date) {
        ArrayList<Object> data = new ArrayList<>();
        String firstName = "";
        String lastName = "";
        String department = "";
        ArrayList<Object> time = new ArrayList<>();

        DefaultTableModel tableModel = (DefaultTableModel) tblAbsensi3.getModel();
        try {
            myStmt = myConn.prepareStatement("select * from data_karyawan where nik=?;");
            myStmt.setString(1, nik);
            myRs = myStmt.executeQuery();
            data.add(nik);
            while (myRs.next()) {
                firstName = myRs.getString("first_name");
                lastName = myRs.getString("last_name");
                department = myRs.getString("Department");
            }
            data.add(firstName);
            data.add(lastName);
            data.add(department);
            data.add(date);

            myStmt = myConn.prepareStatement("select time(date) as time from data_absensi where"
                    + " nik=? and date(date)=?;");

            myStmt.setString(1, nik);
            myStmt.setString(2, date);
            // Execute statement
            myRs = myStmt.executeQuery();
            // Process result set
            int i = 1;
            while (myRs.next()) {
                data.add(myRs.getString("time"));
                i++;
            }
            for (int x = i; x <= 8; x++) {
                data.add("");
            }

            DateTime d1 = formatter.parseDateTime((String) data.get(5));
            DateTime d2 = formatter.parseDateTime((String) data.get(3 + i));

            Period period = new Period(d1, d2);
            data.add(String.format("%02d:%02d:%02d", period.getHours(), period.getMinutes(), period.getSeconds()));

            tableModel.addRow(data.toArray());
        } catch (SQLException ex) {
            showError("sql Exception");
        }
    }

    private void loadNikFromDatabase(String date) {
        try {
            nimList.clear();

            myStmt = myConn.prepareStatement("select nik from data_absensi where date(date)=? group by nik;");
            myStmt.setString(1, date);

            // Execute statement
            myRs = myStmt.executeQuery();
            // Process result set
            while (myRs.next()) {
                nimList.add(myRs.getString("nik"));
            }
        } catch (SQLException ex) {
            showError("sql Exception");
        }

    }

    private class LoadWorker extends SwingWorker<Boolean, Void> {

        @Override
        protected void done() {
            try {
                if (get() != null) {
                    pgbMain.setValue(100);
                    pgbMain2.setValue(100);
                    pgbMain3.setValue(100);

                    disabledButtonWhenLoading(true);

                }
            } catch (InterruptedException | ExecutionException ex) {
                Logger.getLogger(FrmMain.class.getName()).log(Level.SEVERE, null, ex);
            }
        }

        @Override
        protected Boolean doInBackground() throws SQLException, InterruptedException {

            pgbMain.setValue(0);
            pgbMain2.setValue(0);
            pgbMain3.setValue(0);

            disabledButtonWhenLoading(false);
            openAndSaveToDatabase();
            return true;
        }
    }

    private void disabledButtonWhenLoading(boolean b) {
        jMenuItem1.setEnabled(b);
        btnLoadData.setEnabled(b);
        btnLoadData2.setEnabled(b);
        btnLoadData3.setEnabled(b);
        btnExportToExcel.setEnabled(b);
        btnExportToExcel2.setEnabled(b);
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        fchOpen = new javax.swing.JFileChooser();
        fchSave = new javax.swing.JFileChooser();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        jPanel1 = new javax.swing.JPanel();
        lblPath = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tblAbsensi = new javax.swing.JTable();
        dateChooserCombo = new datechooser.beans.DateChooserCombo();
        btnLoadData = new javax.swing.JButton();
        btnExportToExcel = new javax.swing.JButton();
        pgbMain = new javax.swing.JProgressBar();
        lblSheetMain1 = new javax.swing.JLabel();
        jPanel2 = new javax.swing.JPanel();
        dateChooserCombo2 = new datechooser.beans.DateChooserCombo();
        lblPath2 = new javax.swing.JLabel();
        btnLoadData2 = new javax.swing.JButton();
        jScrollPane2 = new javax.swing.JScrollPane();
        tblAbsensi2 = new javax.swing.JTable();
        btnExportToExcel2 = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        dateChooserCombo3 = new datechooser.beans.DateChooserCombo();
        pgbMain2 = new javax.swing.JProgressBar();
        lblSheetMain2 = new javax.swing.JLabel();
        jPanel3 = new javax.swing.JPanel();
        dateChooserCombo4 = new datechooser.beans.DateChooserCombo();
        btnLoadData3 = new javax.swing.JButton();
        jScrollPane3 = new javax.swing.JScrollPane();
        tblAbsensi3 = new javax.swing.JTable();
        btnExportToExcel1 = new javax.swing.JButton();
        lblPath3 = new javax.swing.JLabel();
        pgbMain3 = new javax.swing.JProgressBar();
        dateChooserCombo5 = new datechooser.beans.DateChooserCombo();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        lblSheetMain3 = new javax.swing.JLabel();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenuItem1 = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Absensi");

        lblPath.setText("xls file:");

        tblAbsensi.setAutoCreateRowSorter(true);
        tblAbsensi.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "NIK", "First name", "Last name", "Department", "Date", "Sign in", "Sign out", "Working hours", "telat", "full working hours"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.Boolean.class, java.lang.Boolean.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false, false, false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jScrollPane1.setViewportView(tblAbsensi);
        if (tblAbsensi.getColumnModel().getColumnCount() > 0) {
            tblAbsensi.getColumnModel().getColumn(1).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(2).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(3).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(4).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(5).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(6).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(7).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(8).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(9).setResizable(false);
        }

        dateChooserCombo.setCalendarPreferredSize(new java.awt.Dimension(360, 300));

        btnLoadData.setText("Load data");
        btnLoadData.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnLoadDataActionPerformed(evt);
            }
        });

        btnExportToExcel.setText("Export to excel");
        btnExportToExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExportToExcelActionPerformed(evt);
            }
        });

        pgbMain.setBackground(new java.awt.Color(204, 204, 204));
        pgbMain.setForeground(new java.awt.Color(0, 0, 0));
        pgbMain.setStringPainted(true);

        lblSheetMain1.setText(" ");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(lblPath, javax.swing.GroupLayout.DEFAULT_SIZE, 1461, Short.MAX_VALUE)
                    .addComponent(jScrollPane1)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(btnExportToExcel, javax.swing.GroupLayout.PREFERRED_SIZE, 161, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(dateChooserCombo, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(38, 38, 38)
                                .addComponent(btnLoadData)
                                .addGap(18, 18, 18)
                                .addComponent(pgbMain, javax.swing.GroupLayout.PREFERRED_SIZE, 305, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(lblSheetMain1, javax.swing.GroupLayout.PREFERRED_SIZE, 88, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addComponent(lblPath)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(btnLoadData)
                    .addComponent(dateChooserCombo, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(pgbMain, javax.swing.GroupLayout.PREFERRED_SIZE, 27, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(lblSheetMain1)))
                .addGap(9, 9, 9)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 284, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(26, 26, 26)
                .addComponent(btnExportToExcel, javax.swing.GroupLayout.PREFERRED_SIZE, 52, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 207, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("Main", jPanel1);

        dateChooserCombo2.setCalendarPreferredSize(new java.awt.Dimension(360, 300));

        lblPath2.setText("xls file:");

        btnLoadData2.setText("Load data");
        btnLoadData2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnLoadData2ActionPerformed(evt);
            }
        });

        tblAbsensi2.setAutoCreateRowSorter(true);
        tblAbsensi2.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "NIK", "First name", "Last name", "Department", "Date", "Sign in", "Sign out", "Working hours", "telat", "full working hours"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.Boolean.class, java.lang.Boolean.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false, false, false, true
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jScrollPane2.setViewportView(tblAbsensi2);
        if (tblAbsensi2.getColumnModel().getColumnCount() > 0) {
            tblAbsensi2.getColumnModel().getColumn(0).setResizable(false);
            tblAbsensi2.getColumnModel().getColumn(1).setResizable(false);
            tblAbsensi2.getColumnModel().getColumn(2).setResizable(false);
            tblAbsensi2.getColumnModel().getColumn(3).setResizable(false);
            tblAbsensi2.getColumnModel().getColumn(4).setResizable(false);
            tblAbsensi2.getColumnModel().getColumn(5).setResizable(false);
            tblAbsensi2.getColumnModel().getColumn(6).setResizable(false);
            tblAbsensi2.getColumnModel().getColumn(7).setResizable(false);
        }

        btnExportToExcel2.setText("Export to excel");
        btnExportToExcel2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExportToExcel2ActionPerformed(evt);
            }
        });

        jLabel1.setText("From:");

        jLabel2.setText("to:");

        dateChooserCombo3.setCalendarPreferredSize(new java.awt.Dimension(360, 300));

        pgbMain2.setForeground(new java.awt.Color(0, 0, 0));
        pgbMain2.setStringPainted(true);

        lblSheetMain2.setText(" ");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addComponent(dateChooserCombo2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(121, 121, 121)
                                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel2)
                                    .addComponent(dateChooserCombo3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addComponent(btnLoadData2, javax.swing.GroupLayout.PREFERRED_SIZE, 155, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(pgbMain2, javax.swing.GroupLayout.PREFERRED_SIZE, 305, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(lblSheetMain2, javax.swing.GroupLayout.PREFERRED_SIZE, 88, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(887, 887, Short.MAX_VALUE))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addComponent(btnExportToExcel2, javax.swing.GroupLayout.PREFERRED_SIZE, 161, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addComponent(jLabel1)
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addComponent(lblPath2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jScrollPane2))
                        .addContainerGap())))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addComponent(lblPath2)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jLabel1)
                    .addComponent(jLabel2))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(dateChooserCombo3, javax.swing.GroupLayout.DEFAULT_SIZE, 25, Short.MAX_VALUE)
                    .addComponent(dateChooserCombo2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(8, 8, 8)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(btnLoadData2)
                    .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(pgbMain2, javax.swing.GroupLayout.PREFERRED_SIZE, 27, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(lblSheetMain2)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 284, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(btnExportToExcel2, javax.swing.GroupLayout.PREFERRED_SIZE, 52, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(165, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("Main 2", jPanel2);

        dateChooserCombo4.setCalendarPreferredSize(new java.awt.Dimension(360, 300));

        btnLoadData3.setText("Load data");
        btnLoadData3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnLoadData3ActionPerformed(evt);
            }
        });

        tblAbsensi3.setAutoCreateRowSorter(true);
        tblAbsensi3.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "NIK", "First name", "Last name", "Department", "Date", "clock 1", "clock 2", "clock 3", "clock 4", "clock 5", "clock 6", "clock 7", "clock 8", "Total hours"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false, false, false, false, false, false, false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jScrollPane3.setViewportView(tblAbsensi3);
        if (tblAbsensi3.getColumnModel().getColumnCount() > 0) {
            tblAbsensi3.getColumnModel().getColumn(1).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(2).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(3).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(4).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(5).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(6).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(7).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(8).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(9).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(10).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(11).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(12).setResizable(false);
            tblAbsensi3.getColumnModel().getColumn(13).setResizable(false);
        }

        btnExportToExcel1.setText("Export to excel");
        btnExportToExcel1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExportToExcel1ActionPerformed(evt);
            }
        });

        lblPath3.setText("xls file:");

        pgbMain3.setForeground(new java.awt.Color(0, 0, 0));
        pgbMain3.setStringPainted(true);

        dateChooserCombo5.setCalendarPreferredSize(new java.awt.Dimension(360, 300));

        jLabel3.setText("From:");

        jLabel4.setText("to:");

        lblSheetMain3.setText(" ");

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane3, javax.swing.GroupLayout.DEFAULT_SIZE, 1471, Short.MAX_VALUE)
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(lblPath3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addComponent(btnLoadData3, javax.swing.GroupLayout.PREFERRED_SIZE, 155, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(29, 29, 29)
                                .addComponent(pgbMain3, javax.swing.GroupLayout.PREFERRED_SIZE, 303, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(lblSheetMain3, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 0, Short.MAX_VALUE)))
                        .addContainerGap())
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(btnExportToExcel1, javax.swing.GroupLayout.PREFERRED_SIZE, 161, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(dateChooserCombo4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jLabel3))
                                .addGap(101, 101, 101)
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel4)
                                    .addComponent(dateChooserCombo5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))))
                        .addGap(0, 0, Short.MAX_VALUE))))
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addComponent(lblPath3)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(jLabel4))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(dateChooserCombo4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(dateChooserCombo5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnLoadData3)
                    .addComponent(pgbMain3, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(lblSheetMain3))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 284, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(26, 26, 26)
                .addComponent(btnExportToExcel1, javax.swing.GroupLayout.PREFERRED_SIZE, 52, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(144, 144, 144))
        );

        jTabbedPane1.addTab("Main 3", jPanel3);

        jMenu1.setText("File");

        jMenuItem1.setText("import xls");
        jMenuItem1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem1ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem1);

        jMenuBar1.add(jMenu1);

        setJMenuBar(jMenuBar1);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jTabbedPane1)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jTabbedPane1)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jMenuItem1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem1ActionPerformed
        if (fchOpen.showOpenDialog(this) != 1) {
            path = fchOpen.getSelectedFile().getAbsolutePath();
            worker = new LoadWorker();
            worker.execute();
        }
    }//GEN-LAST:event_jMenuItem1ActionPerformed

    private void btnLoadDataActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLoadDataActionPerformed
        DefaultTableModel tableModel = (DefaultTableModel) tblAbsensi.getModel();

        Date date = dateChooserCombo.getSelectedDate().getTime();
        SimpleDateFormat format1 = new SimpleDateFormat("yyyy-MM-dd");
        String formatted = format1.format(date.getTime());

        loadNikFromDatabase(formatted);

        for (int i = tblAbsensi.getRowCount() - 1; i >= 0; i--) {
            tableModel.removeRow(i);
        }
        nimList.forEach((a) -> {
            loadResultDataFromDatabase(a, formatted, tblAbsensi);
        });
    }//GEN-LAST:event_btnLoadDataActionPerformed

    private void loadTableToObject(JTable table) {
        objectList = new ArrayList<>();
        Object[] a = {"NIK", "First Name", "Last Name", "Department", "Date", "Sign in", "Sign out", "Working hours", "telat", "full working hours"};
        objectList.add(a);
        DefaultTableModel tableModel = (DefaultTableModel) table.getModel();
        for (int i = 0; i <= table.getRowCount() - 1; i++) {
            Object[] b = {tableModel.getValueAt(i, 0), tableModel.getValueAt(i, 1), tableModel.getValueAt(i, 2),
                tableModel.getValueAt(i, 3), tableModel.getValueAt(i, 4), tableModel.getValueAt(i, 5),
                tableModel.getValueAt(i, 6), tableModel.getValueAt(i, 7),
                Boolean.valueOf(tableModel.getValueAt(i, 8).toString()) ? "TELAT" : "",
                Boolean.valueOf(tableModel.getValueAt(i, 9).toString()) ? "" : "------"};
            objectList.add(b);
        }
    }

    private void loadTableToObject2(JTable table) {
        objectList = new ArrayList<>();
        Object[] a = {"NIK", "First Name", "Last Name", "Department", "Date", "Clock 1", "Clock 2", "Clock 3", "Clock 4",
            "Clock 5", "Clock 6", "Clock 7", "Clock 8", "Total Hours"};
        objectList.add(a);
        DefaultTableModel tableModel = (DefaultTableModel) table.getModel();
        for (int i = 0; i <= table.getRowCount() - 1; i++) {
            Object[] b = {tableModel.getValueAt(i, 0), tableModel.getValueAt(i, 1), tableModel.getValueAt(i, 2),
                tableModel.getValueAt(i, 3), tableModel.getValueAt(i, 4), tableModel.getValueAt(i, 5),
                tableModel.getValueAt(i, 6), tableModel.getValueAt(i, 7), tableModel.getValueAt(i, 8),
                tableModel.getValueAt(i, 9), tableModel.getValueAt(i, 10), tableModel.getValueAt(i, 11),
                tableModel.getValueAt(i, 12), tableModel.getValueAt(i, 13)};
            objectList.add(b);
        }
    }

    private void exportToExcel(JTable table) {
        if (fchOpen.showSaveDialog(this) != 1) {

            // set label path
            path2 = fchOpen.getSelectedFile().getAbsolutePath();

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Report");

            if (table.equals(tblAbsensi3)) {
                loadTableToObject2(table);
            } else {
                loadTableToObject(table);
            }

            int rowNum = 0;

            for (Object[] data : objectList) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (Object field : data) {
                    Cell cell = row.createCell(colNum++);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }
            }

            String pathExcel = String.format("%s.xlsx", path2);
            try {
                FileOutputStream outputStream = new FileOutputStream(pathExcel);
                workbook.write(outputStream);
                workbook.close();
            } catch (FileNotFoundException e) {
                showError("File not found exception");
            } catch (IOException e) {
                showError("IO exception");
            }

        }
    }

    private void showError(String msg) {
        JOptionPane.showMessageDialog(this, msg, "Error", 1);
    }

    private void loadDatawithMultipleDate(JTable table, DateChooserCombo dateA, DateChooserCombo dateB) {
        DefaultTableModel tableModel = (DefaultTableModel) table.getModel();
        for (int i = table.getRowCount() - 1; i >= 0; i--) {
            tableModel.removeRow(i);
        }
        String formattedDate;
        Date d1 = dateA.getSelectedDate().getTime();
        Date d2 = dateB.getSelectedDate().getTime();
        DateTime dt1 = new DateTime(d1);
        DateTime dt2 = new DateTime(d2);
        SimpleDateFormat format1 = new SimpleDateFormat("yyyy-MM-dd");
        for (int i = 0; i <= Days.daysBetween(dt1, dt2).getDays(); i++) {
            formattedDate = format1.format(d1.getTime());
            loadNikFromDatabase(format1.format(d1.getTime()));

            for (String a : nimList) {
                if (table.equals(tblAbsensi2)) {
                    loadResultDataFromDatabase(a, formattedDate, table);
                } else {
                    loadResultDataFromDatabase2(a, formattedDate);
                }
            }

            Calendar c = Calendar.getInstance();
            c.setTime(d1);
            c.add(Calendar.DATE, 1);  // number of days to add
            d1 = c.getTime();
        }
    }

    private void btnExportToExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExportToExcelActionPerformed
        exportToExcel(tblAbsensi);
    }//GEN-LAST:event_btnExportToExcelActionPerformed

    private void btnLoadData2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLoadData2ActionPerformed
        loadDatawithMultipleDate(tblAbsensi2, dateChooserCombo2, dateChooserCombo3);
    }//GEN-LAST:event_btnLoadData2ActionPerformed

    private void btnExportToExcel2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExportToExcel2ActionPerformed
        exportToExcel(tblAbsensi2);
    }//GEN-LAST:event_btnExportToExcel2ActionPerformed

    private void btnLoadData3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLoadData3ActionPerformed
        loadDatawithMultipleDate(tblAbsensi3, dateChooserCombo4, dateChooserCombo5);
    }//GEN-LAST:event_btnLoadData3ActionPerformed

    private void btnExportToExcel1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExportToExcel1ActionPerformed
        exportToExcel(tblAbsensi3);
    }//GEN-LAST:event_btnExportToExcel1ActionPerformed

    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(FrmMain.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> {
            new FrmMain().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnExportToExcel;
    private javax.swing.JButton btnExportToExcel1;
    private javax.swing.JButton btnExportToExcel2;
    private javax.swing.JButton btnLoadData;
    private javax.swing.JButton btnLoadData2;
    private javax.swing.JButton btnLoadData3;
    private datechooser.beans.DateChooserCombo dateChooserCombo;
    private datechooser.beans.DateChooserCombo dateChooserCombo2;
    private datechooser.beans.DateChooserCombo dateChooserCombo3;
    private datechooser.beans.DateChooserCombo dateChooserCombo4;
    private datechooser.beans.DateChooserCombo dateChooserCombo5;
    private javax.swing.JFileChooser fchOpen;
    private javax.swing.JFileChooser fchSave;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JLabel lblPath;
    private javax.swing.JLabel lblPath2;
    private javax.swing.JLabel lblPath3;
    private javax.swing.JLabel lblSheetMain1;
    private javax.swing.JLabel lblSheetMain2;
    private javax.swing.JLabel lblSheetMain3;
    private javax.swing.JProgressBar pgbMain;
    private javax.swing.JProgressBar pgbMain2;
    private javax.swing.JProgressBar pgbMain3;
    private javax.swing.JTable tblAbsensi;
    private javax.swing.JTable tblAbsensi2;
    private javax.swing.JTable tblAbsensi3;
    // End of variables declaration//GEN-END:variables
}
