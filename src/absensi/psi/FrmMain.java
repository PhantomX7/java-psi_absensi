/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package absensi.psi;

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
import util.DbConn;

/**
 *
 * @author Lab04
 */
public class FrmMain extends javax.swing.JFrame {

    /**
     * Creates new form FrmMain
     */
    private String path = null;//open
    private String path2 = null;//save

    private Employee employee = null;

    private Connection myConn = null;
    private PreparedStatement myStmt = null;
    private ResultSet myRs = null;

    private ArrayList<String> nimList = new ArrayList<>();

    private String formatted;//date
    private String formatted2;

    private ArrayList<Object[]> objectList = null;

    private LoadWorker worker;

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
            XSSFSheet workSheet = workBook.getSheetAt(0);
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
                double percentage = (double) i * ((double) 100 / (list.size() - 1));
                pgbMain.setValue((int) percentage);
                pgbMain2.setValue((int) percentage);
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

        } catch (FileNotFoundException ex) {
            Logger.getLogger(FrmMain.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException | EncryptedDocumentException | SQLException ex) {
            showError("IO/sql exception");
        } catch (ParseException ex) {
            
        }
    }

    private void loadResultDataFromDatabase(String nik) {

        DefaultTableModel tableModel = (DefaultTableModel) tblAbsensi.getModel();
        try {
            myStmt = myConn.prepareStatement("select nik,First_Name,Last_Name,Department,date(date) as date,min(time(date))"
                    + " as sign_in,max(time(date)) as sign_out,timediff((select max(time(date))"
                    + " from full_data where nik=? and date(date)=?),"
                    + "(select min(time(date)) from full_data where nik=? and date(date)=?))"
                    + " as working_hours  from full_data where nik=? and date(date)=? ");

            myStmt.setString(1, nik);
            myStmt.setString(2, formatted);
            myStmt.setString(3, nik);
            myStmt.setString(4, formatted);
            myStmt.setString(5, nik);
            myStmt.setString(6, formatted);
            // Execute statement
            myRs = myStmt.executeQuery();
            // Process result set
            while (myRs.next()) {
                Object data[] = {myRs.getString("nik"), myRs.getString("first_name"), myRs.getString("last_name"), myRs.getString("Department"),
                    myRs.getString("date"), myRs.getString("sign_in"), myRs.getString("sign_out"), myRs.getString("working_hours")};
                tableModel.addRow(data);
            }
        } catch (SQLException ex) {
            showError("sql Exception");
        }
    }

    private void loadResultDataFromDatabase(String nik, String date) {

        DefaultTableModel tableModel = (DefaultTableModel) tblAbsensi2.getModel();
        try {
            myStmt = myConn.prepareStatement("select nik,First_Name,Last_Name,Department,date(date) as date,min(time(date))"
                    + " as sign_in,max(time(date)) as sign_out,timediff((select max(time(date))"
                    + " from full_data where nik=? and date(date)=?),"
                    + "(select min(time(date)) from full_data where nik=? and date(date)=?))"
                    + " as working_hours  from full_data where nik=? and date(date)=? ");

            myStmt.setString(1, nik);
            myStmt.setString(2, date);
            myStmt.setString(3, nik);
            myStmt.setString(4, date);
            myStmt.setString(5, nik);
            myStmt.setString(6, date);
            // Execute statement
            myRs = myStmt.executeQuery();
            // Process result set
            while (myRs.next()) {
                Object data[] = {myRs.getString("nik"), myRs.getString("first_name"), myRs.getString("last_name"), myRs.getString("Department"),
                    myRs.getString("date"), myRs.getString("sign_in"), myRs.getString("sign_out"), myRs.getString("working_hours")};
                tableModel.addRow(data);
            }
        } catch (SQLException ex) {
            showError("sql Exception");
        }
    }

    private void loadNikFromDatabase() {
        try {
            nimList.clear();

            myStmt = myConn.prepareStatement("select nik from data_absensi where date(date)=? group by nik;");
            Date date = dateChooserCombo.getSelectedDate().getTime();
            SimpleDateFormat format1 = new SimpleDateFormat("yyyy-MM-dd");
            formatted = format1.format(date.getTime());
            myStmt.setString(1, formatted);

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

                    jMenuItem1.setEnabled(true);
                    btnLoadData.setEnabled(true);
                    btnLoadData2.setEnabled(true);
                    btnExportToExcel.setEnabled(true);
                    btnExportToExcel2.setEnabled(true);
                }
            } catch (InterruptedException | ExecutionException ex) {
                Logger.getLogger(FrmMain.class.getName()).log(Level.SEVERE, null, ex);
            }
        }

        @Override
        protected Boolean doInBackground() throws SQLException, InterruptedException {

            pgbMain.setValue(0);
            pgbMain2.setValue(0);

            jMenuItem1.setEnabled(false);
            btnLoadData.setEnabled(false);
            btnLoadData2.setEnabled(false);
            btnExportToExcel.setEnabled(false);
            btnExportToExcel2.setEnabled(false);
            openAndSaveToDatabase();
            return true;
        }
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
                "NIK", "First name", "Last name", "Department", "Date", "Sign in", "Sign out", "Working hours"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false, false
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
            tblAbsensi.getColumnModel().getColumn(0).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(1).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(2).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(3).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(4).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(5).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(6).setResizable(false);
            tblAbsensi.getColumnModel().getColumn(7).setResizable(false);
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

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(lblPath, javax.swing.GroupLayout.DEFAULT_SIZE, 1201, Short.MAX_VALUE)
                    .addComponent(jScrollPane1)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(btnExportToExcel, javax.swing.GroupLayout.PREFERRED_SIZE, 161, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(dateChooserCombo, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(38, 38, 38)
                                .addComponent(btnLoadData)
                                .addGap(18, 18, 18)
                                .addComponent(pgbMain, javax.swing.GroupLayout.PREFERRED_SIZE, 305, javax.swing.GroupLayout.PREFERRED_SIZE)))
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
                    .addComponent(pgbMain, javax.swing.GroupLayout.PREFERRED_SIZE, 27, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(9, 9, 9)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 284, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(26, 26, 26)
                .addComponent(btnExportToExcel, javax.swing.GroupLayout.PREFERRED_SIZE, 52, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 149, Short.MAX_VALUE))
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
                "NIK", "First name", "Last name", "Department", "Date", "Sign in", "Sign out", "Working hours"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false, false
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
                                .addComponent(pgbMain2, javax.swing.GroupLayout.PREFERRED_SIZE, 305, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(733, 733, Short.MAX_VALUE))
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
                    .addComponent(pgbMain2, javax.swing.GroupLayout.PREFERRED_SIZE, 27, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 284, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(btnExportToExcel2, javax.swing.GroupLayout.PREFERRED_SIZE, 52, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(107, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("Main 2", jPanel2);

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
        loadNikFromDatabase();
        for (int i = tblAbsensi.getRowCount() - 1; i >= 0; i--) {
            tableModel.removeRow(i);
        }
        nimList.forEach((a) -> {
            loadResultDataFromDatabase(a);
        });
    }//GEN-LAST:event_btnLoadDataActionPerformed

    private void loadTableToObject(JTable table) {
        objectList = new ArrayList<>();
        Object[] a = {"NIK", "First Name", "Last Name", "Department", "Date", "Sign in", "Sign out", "Working hours"};
        objectList.add(a);
        DefaultTableModel tableModel = (DefaultTableModel) table.getModel();
        for (int i = 0; i <= table.getRowCount() - 1; i++) {
            Object[] b = {tableModel.getValueAt(i, 0), tableModel.getValueAt(i, 1), tableModel.getValueAt(i, 2),
                tableModel.getValueAt(i, 3), tableModel.getValueAt(i, 4), tableModel.getValueAt(i, 5),
                tableModel.getValueAt(i, 6), tableModel.getValueAt(i, 7)};
            objectList.add(b);
        }
    }

    private void exportToExcel(JTable table) {
        if (fchOpen.showSaveDialog(this) != 1) {

            // set label path
            path2 = fchOpen.getSelectedFile().getAbsolutePath();

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Report");

            loadTableToObject(table);

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

    private void btnExportToExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExportToExcelActionPerformed
        exportToExcel(tblAbsensi);
    }//GEN-LAST:event_btnExportToExcelActionPerformed

    private void btnLoadData2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLoadData2ActionPerformed

        DefaultTableModel tableModel = (DefaultTableModel) tblAbsensi2.getModel();
        for (int i = tblAbsensi2.getRowCount() - 1; i >= 0; i--) {
            tableModel.removeRow(i);
        }
        String formattedDate;
        Date d1 = dateChooserCombo2.getSelectedDate().getTime();
        Date d2 = dateChooserCombo3.getSelectedDate().getTime();
        DateTime dt1 = new DateTime(d1);
        DateTime dt2 = new DateTime(d2);
        for (int i = 0; i <= Days.daysBetween(dt1, dt2).getDays(); i++) {
            SimpleDateFormat format1 = new SimpleDateFormat("yyyy-MM-dd");
            formattedDate = format1.format(d1.getTime());
            loadNikFromDatabase(format1.format(d1.getTime()));

            for (String a : nimList) {
                loadResultDataFromDatabase(a, formattedDate);
            }

            Calendar c = Calendar.getInstance();
            c.setTime(d1);
            c.add(Calendar.DATE, 1);  // number of days to add
            d1 = c.getTime();
            System.out.println(d1);
        }
    }//GEN-LAST:event_btnLoadData2ActionPerformed

    private void btnExportToExcel2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExportToExcel2ActionPerformed
        exportToExcel(tblAbsensi2);
    }//GEN-LAST:event_btnExportToExcel2ActionPerformed

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
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(FrmMain.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(FrmMain.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(FrmMain.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(FrmMain.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> {
            new FrmMain().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnExportToExcel;
    private javax.swing.JButton btnExportToExcel2;
    private javax.swing.JButton btnLoadData;
    private javax.swing.JButton btnLoadData2;
    private datechooser.beans.DateChooserCombo dateChooserCombo;
    private datechooser.beans.DateChooserCombo dateChooserCombo2;
    private datechooser.beans.DateChooserCombo dateChooserCombo3;
    private javax.swing.JFileChooser fchOpen;
    private javax.swing.JFileChooser fchSave;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JLabel lblPath;
    private javax.swing.JLabel lblPath2;
    private javax.swing.JProgressBar pgbMain;
    private javax.swing.JProgressBar pgbMain2;
    private javax.swing.JTable tblAbsensi;
    private javax.swing.JTable tblAbsensi2;
    // End of variables declaration//GEN-END:variables
}
