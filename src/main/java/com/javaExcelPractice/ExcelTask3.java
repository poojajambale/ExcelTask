package com.javaExcelPractice;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Perfect Code
@SuppressWarnings("serial")
public class ExcelTask3 extends JFrame {

	String path1;
	String path2;
	int key1;
	int key2;
	String folderPath;

//	String fileName1;
//    String fileName2;

	private void fetchExcel(String path1, String path2, int keyFile1, int keyFile2, String folderPath) {

		try {

			String firstExcelPath = path1;
			FileInputStream file1 = new FileInputStream(firstExcelPath);
			XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
			XSSFSheet sheet1 = workBook1.getSheetAt(0);

			String secondExcelPath = path2;
			FileInputStream file2 = new FileInputStream(secondExcelPath);
			XSSFWorkbook workBook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet2 = workBook2.getSheetAt(0);

			// workBook1
			int totalNumberOfRowsInExcel1 = sheet1.getLastRowNum();
			int totalNumberOfColumnInExcel1 = sheet1.getRow(0).getLastCellNum();
			XSSFCell cellOfRowKey1;
			XSSFRow rowOfSameKey1;

			// workBook2
			int totalNumberOfRowsInExcel2 = sheet2.getLastRowNum();
			int totalNumberOfColumnInExcel2 = sheet2.getRow(0).getLastCellNum();
			XSSFCell cellOfRowKey2;
			XSSFRow rowOfSameKey2;

			// going to Excel1 key -> row = 1 to last
			for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {
				if (sheet1.getRow(r) == null) {
					continue;
				} else {
					if (sheet1.getRow(r).getCell(keyFile1) == null) {
						continue;
					} else {
						cellOfRowKey1 = sheet1.getRow(r).getCell(keyFile1);
					}
//					System.out.println("cellOfRowKey1:"+cellOfRowKey1);
					// going to Excel2 key -> row = 1 to last
					for (int e = 1; e <= totalNumberOfRowsInExcel2; e++) {
						if (sheet2.getRow(e) == null) {
							continue;
						} else {
							if (sheet2.getRow(e).getCell(keyFile2) == null) {
								continue;
							} else {
								cellOfRowKey2 = sheet2.getRow(e).getCell(keyFile2);
							}
//							System.out.println("cellOfRowKey2:"+cellOfRowKey2);
//					cellOfRowKey2 = sheet2.getRow(e).getCell(key2);
							if (cellOfRowKey1.getNumericCellValue() == cellOfRowKey2.getNumericCellValue()) {
//						System.out.println("SameCells1:" + cellOfRowKey1);
								rowOfSameKey1 = sheet1.getRow(r);
								sheet1.removeRow(rowOfSameKey1);
//						sheet1.removeRowBreak(r);
//						removeRow(sheet1, r);
								break;
							}

						} // else
					} // for
				} // else
			} // for

			String firstExcelPathCopy = path1;
//			String firstExcelPathCopy = "C:\\Users\\SATYASAH\\Downloads\\Capg Bench.xlsx";
			FileInputStream file1Copy = new FileInputStream(firstExcelPathCopy);
			XSSFWorkbook workBook1Copy = new XSSFWorkbook(file1Copy);
			XSSFSheet sheet1Copy = workBook1Copy.getSheetAt(0);
			XSSFCell cellOfRowKey1Copy;

			// going to Excel2 key -> row = 1 to last
			for (int rr = 1; rr <= totalNumberOfRowsInExcel2; rr++) {
				if (sheet2.getRow(rr) == null) {
					continue;
				} else {
					if (sheet2.getRow(rr).getCell(keyFile2) == null) {
						continue;
					} else {
						cellOfRowKey2 = sheet2.getRow(rr).getCell(keyFile2);
					}
					// going to Excel1 key -> row = 1 to last
					for (int e = 1; e <= totalNumberOfRowsInExcel1; e++) {
						if (sheet1Copy.getRow(e) == null) {
							continue;
						} else {
							if (sheet1Copy.getRow(e).getCell(keyFile1) == null) {
								continue;
							} else {
								cellOfRowKey1Copy = sheet1Copy.getRow(e).getCell(keyFile1);
							}
//					cellOfRowKey1Copy = sheet1Copy.getRow(e).getCell(key1);
							if (cellOfRowKey2.getNumericCellValue() == cellOfRowKey1Copy.getNumericCellValue()) {
//						System.out.println("SameCells2:" + cellOfRowKey2);
								rowOfSameKey2 = sheet2.getRow(rr);
								sheet2.removeRow(rowOfSameKey2);
								break;
							}
						} // else
					} // for
				} // else
			} // for

//			String target1Path = folderPath + "\\outputFileWithSpace_1.xlsx";
//			FileOutputStream outputStream1 = new FileOutputStream(target1Path);
//			workBook1.write(outputStream1);
//			
//			String target2Path = folderPath + "\\outputFileWithSpace_2.xlsx";
//			FileOutputStream outputStream2 = new FileOutputStream(target2Path);
//			workBook2.write(outputStream2);

			// Upto here we have to two excel with some null or empty row
			// sheet1 and sheet2 as output only NO new sheet created

//-----------------------------------------------------------------------------------------------------------------			

			// counting null row in EXCEL 1
			int counter = 0;
			for (int r = 0; r <= totalNumberOfRowsInExcel1; r++) {
				if (sheet1.getRow(r) == null) {
					counter++;
				}
			}

			if (counter != 0) {

//			System.out.println("totalNumberOfRows1:" + totalNumberOfRowsInExcel1);
//			System.out.println("counter:" + counter);

				int totalNumberOfRowsOfNewSheet = totalNumberOfRowsInExcel1 - counter;

//			System.out.println("totalNumberOfRowsOfNewSheet1:" + totalNumberOfRowsOfNewSheet);

				// creating new working and adding new rows for excel1
				XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
				XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
				XSSFRow rowCreated = null;
//				XSSFCell cellCreated = null;

				for (int r = 0; r <= totalNumberOfRowsOfNewSheet; r++) {
					rowCreated = sheetCreate1.createRow(r);

					for (int c = 0; c < totalNumberOfColumnInExcel1; c++) {
//						cellCreated = rowCreated.createCell(c);
						rowCreated.createCell(c);
					}
				}

				for (int p = 0, u = 0; p <= totalNumberOfRowsInExcel1; p++) {
					if (sheet1.getRow(p) == null) {
						continue;
					} else {
						rowCreated = sheetCreate1.getRow(u);

						for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {
							if (sheet1.getRow(p).getCell(d) == null) {
								continue;
							} else {
//							if (sheet1.getRow(p).getCell(d).getCellType() == null) {
//								continue;
//							}
								if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
						u++;
					}

				}

				// removed null excel writing
				String target1Path1 = folderPath + "\\outputFileOfFinal_1.xlsx";
				FileOutputStream outputStream11 = new FileOutputStream(target1Path1);
				workBookOutput1.write(outputStream11);
				workBookOutput1.close();

			} // if counter close
			else {
				String target1Path = folderPath + "\\outputFileWithSpace_1.xlsx";
				FileOutputStream outputStream1 = new FileOutputStream(target1Path);
				workBook1.write(outputStream1);
			}

			// counting null row in EXCEL 2
			int counter2 = 0;
			for (int r = 1; r <= totalNumberOfRowsInExcel2; r++) {
				if (sheet2.getRow(r) == null) {
					counter2++;
				}
			}

			if (counter2 != 0) {

//			System.out.println("totalNumberOfRows2:" + totalNumberOfRowsInExcel2);
//			System.out.println("counter2:" + counter2);

				int totalNumberOfRowsOfNewSheet2 = totalNumberOfRowsInExcel2 - counter2;

//			System.out.println("totalNumberOfRowsOfNewSheet2:" + totalNumberOfRowsOfNewSheet2);

				// creating new working and adding new rows for excel2
				XSSFWorkbook workBookOutput2 = new XSSFWorkbook();
				XSSFSheet sheetCreate2 = workBookOutput2.createSheet();
				XSSFRow rowCreated2 = null;
//				XSSFCell cellCreated2 = null;

				for (int r = 0; r <= totalNumberOfRowsOfNewSheet2; r++) {
					rowCreated2 = sheetCreate2.createRow(r);

					for (int c = 0; c < totalNumberOfColumnInExcel2; c++) {
//						cellCreated2 = rowCreated2.createCell(c);
						rowCreated2.createCell(c);
					}
				}

				for (int p = 0, v = 0; p <= totalNumberOfRowsInExcel2; p++) {
					if (sheet2.getRow(p) == null) {
						continue;
					} else {
						rowCreated2 = sheetCreate2.getRow(v);

						for (int d = 0; d < totalNumberOfColumnInExcel2; d++) {
							if (sheet2.getRow(p).getCell(d) == null) {
								continue;
							} else {
//							if (sheet2.getRow(p).getCell(d).getCellType() == null) {
//								continue;
//							}
								if (sheet2.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
						v++;
					}
				}

				// null row removed successfully
				// here we will have to two sheetCreate1 and sheetCreate2

				String target1Path2 = folderPath + "\\outputFileOfFinal_2.xlsx";
				FileOutputStream outputStream22 = new FileOutputStream(target1Path2);
				workBookOutput2.write(outputStream22);
				workBookOutput2.close();

			} // if counter close
			else {
				String target2Path = folderPath + "\\outputFileWithSpace_2.xlsx";
				FileOutputStream outputStream2 = new FileOutputStream(target2Path);
				workBook2.write(outputStream2);
			}

			workBook1Copy.close();
			workBook1.close();
			workBook2.close();

//-----------------------------------------------------------------------------------------------------------

			System.out.println("Done....");

		} catch (Exception e) {
			e.printStackTrace();
		}
	} // end of fetch method

//	swing started
//------------------------------------------------------------------------------------------------------------

//	class field
	private JLabel labelFILE1 = new JLabel("FILE 1 :");
	private JLabel labelFILE2 = new JLabel("FILE 2 :");
	private JLabel labelKEYFILE1 = new JLabel("KEY 1 :");
	private JLabel labelKEYFILE2 = new JLabel("KEY 2 :");
	private JLabel outputFolder = new JLabel("OUTPUT :");
	private JLabel displayFileName1 = new JLabel();
	private JLabel displayFileName2 = new JLabel();
	private JLabel displayOutputFolder = new JLabel();
//	private JLabel emptyLabel = new JLabel();
//	private JLabel emptySpace = new JLabel("--------------------");
//	private JLabel emptySpace2 = new JLabel("--------------------");
	JComboBox<String> headerDrop = new JComboBox<String>();
	JComboBox<String> headerDrop2 = new JComboBox<String>();
	private JButton buttonFile1 = new JButton("openFile1");
	private JButton buttonFile2 = new JButton("openFile2");
	private JButton buttonOutput = new JButton("openFolder");
	private JTextField field = new JTextField(10);
	private JTextField field2 = new JTextField(10);
	private JTextField textField1 = new JTextField(10);
//	private JTextField textField2 = new JTextField(10);
	private JButton buttonENTER = new JButton("ENTER");
	private JButton buttonClear = new JButton("CLEAR");
	File file;
	Desktop desktop = Desktop.getDesktop();

	FileInputStream file1;
	XSSFWorkbook workBook1;
	XSSFSheet sheet1;

	FileInputStream file2;
	XSSFWorkbook workBook2;
	XSSFSheet sheet2;

	public ExcelTask3() {
		super("EXCEL TASK");

		setLayout(new GridBagLayout());
		GridBagConstraints constraints = new GridBagConstraints();
		constraints.anchor = GridBagConstraints.WEST;
		constraints.insets = new Insets(10, 10, 10, 10);

		constraints.gridy = 0;
		constraints.gridx = 0;
		this.add(labelFILE1, constraints);

		constraints.gridy = 0;
		constraints.gridx = 1;
		add(buttonFile1, constraints);

		buttonFile1.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == buttonFile1) {

					JFileChooser fileChooser = new JFileChooser();
					FileNameExtensionFilter fnef = new FileNameExtensionFilter("Excel file (.xlsx)", "xlsx");
					fileChooser.setFileFilter(fnef);
//					fileChooser.setCurrentDirectory(
//							new File("C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\task\\Excel"));

					int response = fileChooser.showOpenDialog(null);

					if (response == JFileChooser.APPROVE_OPTION) {

//						File file = new File(fileChooser.getSelectedFile().getAbsolutePath());
						File file2 = fileChooser.getSelectedFile();

//						System.out.println(file);
//						fileName1 = file2.getName();

						if (file2.getName().length() < 13) {
							displayFileName1.setText(file2.getName());
						} else {
//						displayFileName2.setText(file2.getName().substring(0, 14));
							displayFileName1.setText(file2.getName().substring(0, 14));
						}
//						displayFileName1.setText(file2.getName());
//						displayFileName1.setText(file2.getName().substring(0, 14));
//						textField1.setText(file2.getName().substring(0,4));

						String s = fileChooser.getSelectedFile().getAbsolutePath();

//						System.out.println(s);

						path1 = s;
//						System.out.println(path1);
						try {
							headerDrop.removeAllItems();
							file1 = new FileInputStream(path1);
							workBook1 = new XSSFWorkbook(file1);
							sheet1 = workBook1.getSheetAt(0);
//							int rows = sheet1.getLastRowNum();
							if (sheet1.getRow(0) == null) {
//								System.out.println("Excel is empty");
								JOptionPane.showMessageDialog(ExcelTask3.this, "Excel file 1 is Empty", "Excel",
										JOptionPane.ERROR_MESSAGE);
								path1 = null;
							} else {
								int column = sheet1.getRow(0).getLastCellNum();

								for (int r = 0; r < 1; r++) {
									XSSFRow row = sheet1.getRow(0);
									for (int c = 0; c < column; c++) {
										XSSFCell cell = row.getCell(c);
										// 1
//									System.out.println("cell" + cell);
										headerDrop.addItem("" + cell);
									} // for
								} // for
//							workBook1.close();
//							file1.close();
							}
						} catch (IOException e1) {
							e1.printStackTrace();
						}
					}
				}
			}
		});

		constraints.gridy = 0;
		constraints.gridx = 2;
		add(displayFileName1, constraints);
//		add(textField1, constraints);
		textField1.setEditable(false);

//		headerDrop.setSize(10,7);
		headerDrop.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXX");

		headerDrop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == headerDrop) {
//					headerDrop.setSelectedIndex(0);
//					setExtendedState(JFrame.MAXIMIZED_BOTH);
//					keyName1 = (String) headerDrop.getSelectedItem();
//					System.out.println(keyName1);
					key1 = headerDrop.getSelectedIndex();
//					System.out.println("key1:"+key1);
//				String selectedHeader = headerDrop.getSelectedItem();
				}
			}
		});

//		headerDrop.setEditable(true);
//		headerDrop.setSelectedIndex(0);
//		headerDrop.setForeground(Color.BLUE);
//		headerDrop.setBackground(Color.WHITE);
//		headerDrop.setFont(new Font("Arial", Font.BOLD, 14));
		// And limit the maximum number of items displayed in the drop-down list:
		headerDrop.setMaximumRowCount(5); // scroller

		constraints.gridy = 1;
		constraints.gridx = 0;
		add(labelFILE2, constraints);

		constraints.gridy = 1;
		constraints.gridx = 1;
		add(buttonFile2, constraints);

		constraints.gridy = 1;
		constraints.gridx = 2;
		add(displayFileName2, constraints);

		buttonFile2.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == buttonFile2) {

					JFileChooser fileChooser = new JFileChooser();

					FileNameExtensionFilter fnef = new FileNameExtensionFilter("Excel file (.xlsx)", "xlsx");
					fileChooser.setFileFilter(fnef);

//					fileChooser.setCurrentDirectory(
//							new File("C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\task\\Excel")); // sets

					int response = fileChooser.showOpenDialog(null);

					if (response == JFileChooser.APPROVE_OPTION) {
//						File file = new File(fileChooser.getSelectedFile().getAbsolutePath());
						File file11 = fileChooser.getSelectedFile();
//						fileName2 = file11.getName();
//					    File file12 = file11.getAbsoluteFile();
//					    System.out.println(file11.getPath());
//						String n = file2.getName().substring(0, 4);
						if (file11.getName().length() < 13) {
							displayFileName2.setText(file11.getName());
						} else {
//						displayFileName2.setText(file2.getName().substring(0, 14));
							displayFileName2.setText(file11.getName().substring(0, 14));
						}

						String s = fileChooser.getSelectedFile().getAbsolutePath();
						path2 = s;
//						System.out.println(path2);
						try {
							headerDrop2.removeAllItems();
							file2 = new FileInputStream(path2);
							workBook2 = new XSSFWorkbook(file2);
							sheet2 = workBook2.getSheetAt(0);
//							int rows = sheet2.getLastRowNum();

							if (sheet2.getRow(0) == null) {
//								System.out.println("Excel is empty");
								JOptionPane.showMessageDialog(ExcelTask3.this, "Excel file 2 is Empty", "Excel",
										JOptionPane.ERROR_MESSAGE);
								path2 = null;
							} else {
								int column = sheet2.getRow(0).getLastCellNum();

								for (int r = 0; r < 1; r++) {
									XSSFRow row = sheet2.getRow(0);
									for (int c = 0; c < column; c++) {
										XSSFCell cell = row.getCell(c);
										// 1
//									System.out.println("cell" + cell);
										headerDrop2.addItem("" + cell);
									} // for
								} // for

							}
						} catch (IOException e1) {
							e1.printStackTrace();
						}

					}
				}
			}
		});

		headerDrop2.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXX");
		headerDrop2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == headerDrop2) {
//					keyName2 = (String) headerDrop2.getSelectedItem();
//					System.out.println(headerDrop2.getSelectedItem());
					key2 = headerDrop2.getSelectedIndex();
//					System.out.println("key2:"+key2);
				}
			}
		});

//		headerDrop2.setEditable(true);
//		headerDrop.setSelectedIndex(key1);
//		headerDrop2.setForeground(Color.BLUE);
//		headerDrop2.setBackground(Color.WHITE);
//		headerDrop.setFont(new Font("Arial", Font.BOLD, 14));
		// And limit the maximum number of items displayed in the drop-down list:
		headerDrop2.setMaximumRowCount(5); // scroller

		constraints.gridy = 3;
		constraints.gridx = 0;
		add(labelKEYFILE1, constraints);

		field.setEditable(false);

//		constraints.anchor = GridBagConstraints.CENTER;
		constraints.gridy = 3;
		constraints.gridx = 2;
		add(field, constraints);
//		add(emptySpace, constraints);

		constraints.gridy = 3;
		constraints.gridx = 1;
		add(headerDrop, constraints);

//		String excel1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\List1.xlsx";

		constraints.gridy = 4;
		constraints.gridx = 0;
		add(labelKEYFILE2, constraints);

		field2.setEditable(false);
//		constraints.anchor = GridBagConstraints.CENTER;
		constraints.gridy = 4;
		constraints.gridx = 2;
		add(field2, constraints);
//		add(emptySpace2, constraints);

		constraints.gridy = 4;
		constraints.gridx = 1;
		add(headerDrop2, constraints);

		constraints.gridx = 0;
		constraints.gridy = 5;
		add(outputFolder, constraints);

		constraints.gridx = 1;
		constraints.gridy = 5;
		add(buttonOutput, constraints);

		constraints.gridx = 2;
		constraints.gridy = 5;
		add(displayOutputFolder, constraints);

		buttonOutput.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == buttonOutput) {

					JFileChooser fileChooser = new JFileChooser();
					fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

//					fileChooser.setCurrentDirectory(
//							new File("C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\task\\Excel")); // sets

					int response = fileChooser.showOpenDialog(ExcelTask3.this);

					if (response == JFileChooser.APPROVE_OPTION) {
						file = new File(fileChooser.getSelectedFile().getAbsolutePath());
						File file2 = fileChooser.getSelectedFile();
						displayOutputFolder.setText(file2.getName());
//						System.out.println(file);
						String s = fileChooser.getSelectedFile().getAbsolutePath();
						folderPath = s;
					} else {
						displayOutputFolder.setText("");
					}
				}
			}
		});

		constraints.gridx = 0;
		constraints.gridy = 6;
		constraints.gridwidth = 3;
		constraints.anchor = GridBagConstraints.CENTER;
		add(buttonENTER, constraints);

		buttonENTER.setBackground(Color.cyan);

		buttonENTER.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent event) {
				if (path1 == null && path2 != null) {
					JOptionPane.showMessageDialog(ExcelTask3.this, "Enter File1", "File", JOptionPane.ERROR_MESSAGE);
				} else if (path1 != null && path2 == null) {
					JOptionPane.showMessageDialog(ExcelTask3.this, "Enter File2", "File", JOptionPane.ERROR_MESSAGE);
				} else if (path1 != null && path2 != null && folderPath == null) {
					JOptionPane.showMessageDialog(ExcelTask3.this, "Select Folder", "Folder",
							JOptionPane.ERROR_MESSAGE);
				} else if (path1 == null && path2 == null && folderPath == null) {
					JOptionPane.showMessageDialog(ExcelTask3.this, "Select Files and Folder", "Folder",
							JOptionPane.ERROR_MESSAGE);
				} else if (path1.equals(path2)) {
					JOptionPane.showMessageDialog(ExcelTask3.this, "Both File 1 and File 2 are Same Select other file",
							"File", JOptionPane.ERROR_MESSAGE);
					path2 = null;
				} else {

					// workBook1
					int initial_totalNumberOfRowsInExcel1 = sheet1.getLastRowNum();

					// workBook2
					int intial_totalNumberOfRowsInExcel2 = sheet2.getLastRowNum();

					int count1 = 0;
					for (int i = 1; i <= initial_totalNumberOfRowsInExcel1; i++) {
						if (sheet1.getRow(i) == null) {
							continue;
						} else {
							if (sheet1.getRow(i).getCell(key1) == null) {
								continue;
							} else {
								if (sheet1.getRow(i).getCell(key1).getCellType() == CellType.STRING) {
									JOptionPane.showMessageDialog(ExcelTask3.this,
											"Key1 in Excel1 is not numeric Choose Again", "Excel",
											JOptionPane.INFORMATION_MESSAGE);
									count1++;
									break;
								} else {
									// do nothing
								}
							} // else
						} // else
					} // for

					int count2 = 0;
					for (int j = 1; j <= intial_totalNumberOfRowsInExcel2; j++) {
						if (sheet2.getRow(j) == null) {
							continue;
						} else {
							if (sheet2.getRow(j).getCell(key2) == null) {
								continue;
							} else {
						if (sheet2.getRow(j).getCell(key2).getCellType() == CellType.STRING) {
							JOptionPane.showMessageDialog(ExcelTask3.this, "Key2 in Excel2 is not numeric Choose Again",
									"Excel", JOptionPane.INFORMATION_MESSAGE);
							count2++;
							break;
						} else {
							// do nothing
						}
					}}}

					if (count1 == 0 && count2 == 0) {
						fetchExcel(path1, path2, key1, key2, folderPath);

						JOptionPane.showMessageDialog(ExcelTask3.this, "Excel created", "Excel",
								JOptionPane.PLAIN_MESSAGE);
						
						try {
							desktop.open(file);
						} catch (IOException e) {
							e.printStackTrace();
						}

						int ii = JOptionPane.showConfirmDialog(null,
								"We Have to close this window in order to open newly generated Excel, Beacause these are already open or are in use by javaw.exe",
								"Exit?", JOptionPane.YES_NO_OPTION);
						if (ii == 1) {
							// do nothing
						}
						if (ii == 0) {
							System.exit(0);
						}

						
					}
				}
			}

		});

		constraints.gridx = 2;
		constraints.gridy = 6;
		constraints.gridwidth = 3;
		constraints.anchor = GridBagConstraints.CENTER;

		add(buttonClear, constraints);

		buttonClear.setBackground(Color.red);

		buttonClear.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				path1 = null;
				path2 = null;
				headerDrop.removeAllItems();
				headerDrop2.removeAllItems();
				folderPath = null;
				displayFileName1.setText("");
				displayFileName2.setText("");
				displayOutputFolder.setText("");

//				JOptionPane.showMessageDialog(ExcelTask3.this, "Field's Cleared", "File", JOptionPane.PLAIN_MESSAGE);
			}

		});

		pack();
//		setExtendedState(JFrame.MAXIMIZED_BOTH);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLocationRelativeTo(null);

	}

	public static void main(String[] args) {

		SwingUtilities.invokeLater(new Runnable() {

			public void run() {
				new ExcelTask3().setVisible(true);
			}
		});
	}
}
