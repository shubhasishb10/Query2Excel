import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import javax.swing.ComboBoxModel;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.Unmarshaller;

/**
 * 
 * @author Shubhasish
 */
@SuppressWarnings("serial")
public class NewJFrame extends javax.swing.JFrame {
	public NewJFrame() {
		super("Query2Excel");
		initComponents();
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void initComponents() {

		// jPopupMenu1 = new javax.swing.JPopupMenu();
		// jPopupMenu2 = new javax.swing.JPopupMenu();
		setResizable(false);
		jDialog1 = new javax.swing.JDialog();
		jDialog2 = new javax.swing.JDialog();
		jDialog3 = new javax.swing.JDialog();
		jScrollPane1 = new javax.swing.JScrollPane();
		jTextArea1 = new javax.swing.JTextArea();
		jComboBox1 = new javax.swing.JComboBox();
		jButton1 = new javax.swing.JButton();
		jComboBox2 = new javax.swing.JComboBox();
		jButton2 = new javax.swing.JButton();
		jButton3 = new javax.swing.JButton();
		jButton4 = new JButton();
		jSeparator1 = new javax.swing.JSeparator();

		javax.swing.GroupLayout jDialog1Layout = new javax.swing.GroupLayout(
				jDialog1.getContentPane());
		jDialog1.getContentPane().setLayout(jDialog1Layout);
		jDialog1Layout.setHorizontalGroup(jDialog1Layout.createParallelGroup(
				javax.swing.GroupLayout.Alignment.LEADING).addGap(0, 400,
				Short.MAX_VALUE));
		jDialog1Layout.setVerticalGroup(jDialog1Layout.createParallelGroup(
				javax.swing.GroupLayout.Alignment.LEADING).addGap(0, 300,
				Short.MAX_VALUE));

		javax.swing.GroupLayout jDialog2Layout = new javax.swing.GroupLayout(
				jDialog2.getContentPane());
		jDialog2.getContentPane().setLayout(jDialog2Layout);
		jDialog2Layout.setHorizontalGroup(jDialog2Layout.createParallelGroup(
				javax.swing.GroupLayout.Alignment.LEADING).addGap(0, 400,
				Short.MAX_VALUE));
		jDialog2Layout.setVerticalGroup(jDialog2Layout.createParallelGroup(
				javax.swing.GroupLayout.Alignment.LEADING).addGap(0, 300,
				Short.MAX_VALUE));

		javax.swing.GroupLayout jDialog3Layout = new javax.swing.GroupLayout(
				jDialog3.getContentPane());
		jDialog3.getContentPane().setLayout(jDialog3Layout);
		jDialog3Layout.setHorizontalGroup(jDialog3Layout.createParallelGroup(
				javax.swing.GroupLayout.Alignment.LEADING).addGap(0, 400,
				Short.MAX_VALUE));
		jDialog3Layout.setVerticalGroup(jDialog3Layout.createParallelGroup(
				javax.swing.GroupLayout.Alignment.LEADING).addGap(0, 300,
				Short.MAX_VALUE));

		setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

		jTextArea1.setBackground(new java.awt.Color(204, 255, 255));
		jTextArea1.setColumns(10);
		jTextArea1.setFont(new java.awt.Font("Calibri", 1, 14)); // NOI18N
		jTextArea1.setForeground(new java.awt.Color(51, 0, 51));
		jTextArea1.setLineWrap(true);
		jTextArea1.setRows(5);
		jTextArea1.setWrapStyleWord(true);
		jTextArea1.setSelectedTextColor(new java.awt.Color(0, 102, 102));
		jTextArea1.setSelectionColor(new java.awt.Color(255, 204, 204));
		jScrollPane1.setViewportView(jTextArea1);

		jComboBox1.setBackground(new java.awt.Color(51, 255, 255));
		jComboBox1.setFont(new java.awt.Font("Verdana", 1, 12)); // NOI18N
		jComboBox1.setForeground(new java.awt.Color(51, 0, 153));
		jComboBox1.setModel(getServerDropDownValue());
		jComboBox1.setToolTipText("Select Server");

		jButton1.setBackground(new java.awt.Color(0, 102, 102));
		jButton1.setFont(new java.awt.Font("Verdana", 0, 14)); // NOI18N
		jButton1.setText("V");
		jButton1.addActionListener(new java.awt.event.ActionListener() {
			public void actionPerformed(java.awt.event.ActionEvent evt) {
				jButton1ActionPerformed(evt);
			}
		});

		jComboBox2.setBackground(new java.awt.Color(51, 255, 255));
		jComboBox2.setFont(new java.awt.Font("Verdana", 1, 12)); // NOI18N
		jComboBox2.setForeground(new java.awt.Color(51, 0, 51));

		jButton2.setBackground(new java.awt.Color(0, 153, 153));
		jButton2.setFont(new java.awt.Font("Verdana", 1, 12)); // NOI18N
		jButton2.setForeground(new java.awt.Color(51, 0, 51));
		jButton2.setText("Generate Excel");
		jButton2.setBorderPainted(false);
		jButton2.addActionListener(new java.awt.event.ActionListener() {
			public void actionPerformed(java.awt.event.ActionEvent evt) {
				jButton2ActionPerformed(evt);
			}
		});

		jButton3.setBackground(new java.awt.Color(0, 153, 153));
		jButton3.setFont(new java.awt.Font("Verdana", 1, 12)); // NOI18N
		jButton3.setForeground(new java.awt.Color(51, 51, 51));
		jButton3.setText("Admin Menu");
		jButton3.addMouseListener(new java.awt.event.MouseAdapter() {
			public void mouseClicked(java.awt.event.MouseEvent evt) {
				jButton3MouseClicked(evt);
			}
		});

		jButton4.setBackground(new java.awt.Color(0, 153, 153));
		jButton4.setFont(new java.awt.Font("Verdana", 1, 12)); // NOI18N
		jButton4.setForeground(new java.awt.Color(51, 51, 51));
		jButton4.setText("Add Server");
		jButton4.addMouseListener(new java.awt.event.MouseAdapter() {
			public void mouseClicked(java.awt.event.MouseEvent evt) {
				jButton4MouseClicked(evt);
			}
		});

		jSeparator1.setBackground(new java.awt.Color(102, 0, 153));

		javax.swing.GroupLayout layout = new javax.swing.GroupLayout(
				getContentPane());
		getContentPane().setLayout(layout);
		layout.setHorizontalGroup(layout
				.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
				.addGroup(
						layout.createSequentialGroup()
								.addContainerGap()
								.addComponent(jScrollPane1,
										javax.swing.GroupLayout.PREFERRED_SIZE,
										617,
										javax.swing.GroupLayout.PREFERRED_SIZE)
								.addGap(7, 7, 7)
								.addComponent(jSeparator1,
										javax.swing.GroupLayout.PREFERRED_SIZE,
										javax.swing.GroupLayout.DEFAULT_SIZE,
										javax.swing.GroupLayout.PREFERRED_SIZE)
								.addGroup(
										layout.createParallelGroup(
												javax.swing.GroupLayout.Alignment.LEADING)
												.addGroup(
														layout.createSequentialGroup()
																.addGap(74, 74,
																		74)
																.addComponent(
																		jButton3,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		142,
																		javax.swing.GroupLayout.PREFERRED_SIZE))
												.addGroup(
														layout.createSequentialGroup()
																.addGap(74, 74,
																		74)
																.addComponent(
																		jButton4,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		142,
																		javax.swing.GroupLayout.PREFERRED_SIZE))
												.addGroup(
														layout.createSequentialGroup()
																.addPreferredGap(
																		javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
																.addComponent(
																		jComboBox2,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		252,
																		javax.swing.GroupLayout.PREFERRED_SIZE))
												.addGroup(
														layout.createSequentialGroup()
																.addGap(117,
																		117,
																		117)
																.addComponent(
																		jButton1))
												.addGroup(
														layout.createSequentialGroup()
																.addGap(18, 18,
																		18)
																.addComponent(
																		jComboBox1,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		252,
																		javax.swing.GroupLayout.PREFERRED_SIZE))
												.addGroup(
														layout.createSequentialGroup()
																.addGap(69, 69,
																		69)
																.addComponent(
																		jButton2,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		142,
																		javax.swing.GroupLayout.PREFERRED_SIZE)))
								.addContainerGap(18, Short.MAX_VALUE)));
		layout.setVerticalGroup(layout
				.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
				.addGroup(
						layout.createSequentialGroup()
								.addContainerGap()
								.addGroup(
										layout.createParallelGroup(
												javax.swing.GroupLayout.Alignment.LEADING)
												.addGroup(
														layout.createSequentialGroup()
																.addGroup(
																		layout.createParallelGroup(
																				javax.swing.GroupLayout.Alignment.TRAILING,
																				false)
																				.addComponent(
																						jScrollPane1,
																						javax.swing.GroupLayout.Alignment.LEADING,
																						javax.swing.GroupLayout.DEFAULT_SIZE,
																						484,
																						Short.MAX_VALUE)
																				.addComponent(
																						jSeparator1,
																						javax.swing.GroupLayout.Alignment.LEADING))
																.addContainerGap(
																		javax.swing.GroupLayout.DEFAULT_SIZE,
																		Short.MAX_VALUE))
												.addGroup(
														layout.createSequentialGroup()
																.addComponent(
																		jButton3,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		43,
																		javax.swing.GroupLayout.PREFERRED_SIZE)
																.addGap(10, 10,
																		10)
																.addComponent(
																		jButton4,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		43,
																		javax.swing.GroupLayout.PREFERRED_SIZE)
																.addGap(10, 70,
																		80)
																.addComponent(
																		jComboBox1,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		36,
																		javax.swing.GroupLayout.PREFERRED_SIZE)
																.addGap(39, 39,
																		39)
																.addComponent(
																		jButton1,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		36,
																		javax.swing.GroupLayout.PREFERRED_SIZE)
																.addGap(37, 37,
																		37)
																.addComponent(
																		jComboBox2,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		38,
																		javax.swing.GroupLayout.PREFERRED_SIZE)
																.addPreferredGap(
																		javax.swing.LayoutStyle.ComponentPlacement.RELATED,
																		79,
																		Short.MAX_VALUE)
																.addComponent(
																		jButton2,
																		javax.swing.GroupLayout.PREFERRED_SIZE,
																		53,
																		javax.swing.GroupLayout.PREFERRED_SIZE)
																.addGap(10, 10,
																		10)))));

		pack();
	}

	@SuppressWarnings("unchecked")
	private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {
		Connection connection = null;
		Statement stmt = null;
		List<String> databasesList = new ArrayList<>();
		try {
			connection = getConnection();
			System.err.println(jComboBox1.getSelectedItem());
			System.err.println(jComboBox1.getSelectedIndex());
			stmt = connection.createStatement();
			String query = "select name from master.dbo.sysdatabases";
			ResultSet rs = stmt.executeQuery(query);
			ResultSetMetaData rsMeta = rs.getMetaData();
			System.out.println(rsMeta.getColumnName(1));
			while (rs.next()) {
				String tempStr = String.valueOf(rs.getObject(1));
				System.out.println(tempStr);
				databasesList.add(tempStr);
			}
			jComboBox2.setModel(new DefaultComboBoxModel<>(databasesList
					.toArray()));
			// select name from master.dbo.sysdatabases
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {
		String query = null;
		Connection con = null;
		Statement stmt = null;
		List<List<String>> wholeList = null;
		List<String> rowList = null;
		try {
			wholeList = new ArrayList<>();
			if (!"".equals(jTextArea1.getText())
					|| jTextArea1.getText() != null) {
				query = "use " + jComboBox2.getSelectedItem().toString() + "\n";
				query += jTextArea1.getText();
				con = getConnection();
				stmt = con.createStatement();
				ResultSet rs = stmt.executeQuery(query);
				rowList = new ArrayList<>();
				for (int i = 0; i < rs.getMetaData().getColumnCount(); i++) {
					rowList.add(rs.getMetaData().getColumnName(i + 1));
				}
				wholeList.add(rowList);
				while (rs.next()) {
					rowList = new ArrayList<>();
					for (int i = 0; i < rs.getMetaData().getColumnCount(); i++) {
						String data = String.valueOf(rs.getObject(i + 1));
						rowList.add(data);
					}
					wholeList.add(rowList);
				}

				new ExcelGenerator().generateExcelDocument(wholeList);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void jButton3MouseClicked(java.awt.event.MouseEvent evt) {
		try {
			AdminForm ad = new AdminForm();
			ad.setDefaultCloseOperation(WIDTH);
			ad.setVisible(rootPaneCheckingEnabled);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void jButton4MouseClicked(java.awt.event.MouseEvent evt) {
		try {
			AddServer adSv = new AddServer();
			adSv.setDefaultCloseOperation(WIDTH);
			adSv.setVisible(rootPaneCheckingEnabled);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String args[]) {

		try {
			for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager
					.getInstalledLookAndFeels()) {
				if ("Nimbus".equals(info.getName())) {
					javax.swing.UIManager.setLookAndFeel(info.getClassName());
					break;
				}
			}
		} catch (ClassNotFoundException ex) {
			java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(
					java.util.logging.Level.SEVERE, null, ex);
		} catch (InstantiationException ex) {
			java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(
					java.util.logging.Level.SEVERE, null, ex);
		} catch (IllegalAccessException ex) {
			java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(
					java.util.logging.Level.SEVERE, null, ex);
		} catch (javax.swing.UnsupportedLookAndFeelException ex) {
			java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(
					java.util.logging.Level.SEVERE, null, ex);
		} catch (Exception e) {
			e.printStackTrace();
		}

		java.awt.EventQueue.invokeLater(new Runnable() {
			public void run() {
				new NewJFrame().setVisible(true);
			}
		});
	}

	@SuppressWarnings("rawtypes")
	public ComboBoxModel getServerDropDownValue() {
		List<Server> servers = null;
		try {
			File serverListFile = new File("src/newXMLDocument.xml");
			JAXBContext context = JAXBContext.newInstance(Servers.class);
			Unmarshaller unmarsheller = context.createUnmarshaller();
			Servers svs = (Servers) unmarsheller.unmarshal(serverListFile);
			servers = svs.getServers();
		} catch (Exception e) {
			e.printStackTrace();
		}

		return new DefaultComboBoxModel<>(servers.toArray());
	}

	public Connection getConnection() {
		Connection connection = null;
		try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			String conURL = "jdbc:sqlserver://"
					+ jComboBox1.getSelectedItem().toString().trim() + ":1433;";
			connection = DriverManager.getConnection(conURL, "v2020", "v2020");
		} catch (Exception e) {
			e.printStackTrace();
		}
		return connection;
	}
	
	@SuppressWarnings("unchecked")
	public void refreshServerList() {
		try {
			jComboBox1.setModel(getServerDropDownValue());
			
		}catch(Exception e) {
			e.printStackTrace();
		}
	}

	private javax.swing.JButton jButton1;
	private javax.swing.JButton jButton2;
	private javax.swing.JButton jButton3;
	private javax.swing.JButton jButton4;
	@SuppressWarnings("rawtypes")
	private javax.swing.JComboBox jComboBox1;
	@SuppressWarnings("rawtypes")
	private javax.swing.JComboBox jComboBox2;
	private javax.swing.JDialog jDialog1;
	private javax.swing.JDialog jDialog2;
	private javax.swing.JDialog jDialog3;
	// private javax.swing.JPopupMenu jPopupMenu1;
	// private javax.swing.JPopupMenu jPopupMenu2;
	private javax.swing.JScrollPane jScrollPane1;
	private javax.swing.JSeparator jSeparator1;
	private javax.swing.JTextArea jTextArea1;
}
