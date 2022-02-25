/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/*
 * JPrincipal.java
 *
 * Created on 10-05-2010, 12:41:39 PM
 */

package jmycode;

import java.sql.*;

/**
 *
 * @author Administrador
 */
public class JPrincipal extends javax.swing.JFrame {

    /** Creates new form JPrincipal */
    public JPrincipal()
    {
        initComponents();
        Conexion();
    }

    public static void Conexion()
    {
        try
        {
            Class.forName("org.sqlite.JDBC");
            Connection conn =
            DriverManager.getConnection("jdbc:sqlite:mycode.exe");
            Statement stat = conn.createStatement();
            stat.executeUpdate("drop table if exists people;");
            stat.executeUpdate("create table people (name, occupation);");
            PreparedStatement prep = conn.prepareStatement(
              "insert into people values (?, ?);");

            prep.setString(1, "Gandhi");
            prep.setString(2, "politics");
            prep.addBatch();
            prep.setString(1, "Turing");
            prep.setString(2, "computers");
            prep.addBatch();
            prep.setString(1, "Wittgenstein");
            prep.setString(2, "smartypants");
            prep.addBatch();

            conn.setAutoCommit(false);
            prep.executeBatch();
            conn.setAutoCommit(true);

            ResultSet rs = stat.executeQuery("select * from people;");
            while (rs.next()) {
              System.out.println("name = " + rs.getString("name"));
              System.out.println("job = " + rs.getString("occupation"));
            }
            rs.close();
            conn.close();
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
  }


    /** This method is called from within the constructor to
     * initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is
     * always regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        getContentPane().setLayout(null);

        jLabel1.setText("label");
        getContentPane().add(jLabel1);
        jLabel1.setBounds(20, 20, 22, 14);

        java.awt.Dimension screenSize = java.awt.Toolkit.getDefaultToolkit().getScreenSize();
        setBounds((screenSize.width-593)/2, (screenSize.height-426)/2, 593, 426);
    }// </editor-fold>//GEN-END:initComponents

    /**
    * @param args the command line arguments
    */
    public static void main(String args[]) {
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new JPrincipal().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JLabel jLabel1;
    // End of variables declaration//GEN-END:variables

}
