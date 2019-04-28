package com.wiwj.cbs;

import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.awt.Frame;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

import com.alibaba.fastjson.JSON;

/**
 * 选择两个要比较的excel文件
 * 
 */
public class CompareExcelFileSelectDialog 
	extends JDialog 
	implements ActionListener {

	private static final long serialVersionUID = 1L;
	// 旧excel路径
	private String oldExcelFilePath;
	// 新excel路径
	private String newExcelFilePath;

	private JTextField oldExcelFileInput;
	private JTextField newExcelFileInput;

	/**
	 * 构造方法，必须指明owner
	 * 
	 * @param owner
	 */
	public CompareExcelFileSelectDialog(Frame owner) {
		super(owner, "BOM文件比较", true);
		initGUI();
	}

	/**
	 * 初始化GUI
	 */
	private void initGUI() {
		// 设定大小和位置
		this.setSize(600, 280);
		this.setLocation(450, 200);
		// 添加组件
		this.setLayout(new BorderLayout());
		FlowLayout northLayout = new FlowLayout(FlowLayout.LEFT);
		JPanel northPanel = new JPanel(northLayout);
		JLabel titleLabel = new JLabel("需要比较的文件");
		northPanel.add(titleLabel);
		this.add(northPanel, BorderLayout.NORTH);
		// 中间面板
		JPanel centerPanel = new JPanel(new GridLayout(2, 1));
		// 选择旧的文件组件
		JPanel centerUpperPanel = new JPanel();
		JLabel oldFolderLabel = new JLabel("旧BOM:");
		oldExcelFileInput = new JTextField(oldExcelFilePath, 40);
		JButton oldFolderSelectButton = new JButton("选择");
		oldFolderSelectButton.setActionCommand("oldFolderSelect");
		oldFolderSelectButton.addActionListener(this);
		centerUpperPanel.add(oldFolderLabel);
		centerUpperPanel.add(oldExcelFileInput);
		centerUpperPanel.add(oldFolderSelectButton);
		// 选择新的文件组件
		JPanel centerLowerPanel = new JPanel();
		JLabel newFolderLabel = new JLabel("新BOM:");
		newExcelFileInput = new JTextField(newExcelFilePath, 40);
		JButton newFolderSelectButton = new JButton("选择");
		newFolderSelectButton.setActionCommand("newFolderSelect");
		newFolderSelectButton.addActionListener(this);
		centerLowerPanel.add(newFolderLabel);
		centerLowerPanel.add(newExcelFileInput);
		centerLowerPanel.add(newFolderSelectButton);
		// 填充中间面板
		centerPanel.add(centerUpperPanel);
		centerPanel.add(centerLowerPanel);
		this.add(centerPanel, BorderLayout.CENTER);
		// 添加确定按钮
		FlowLayout southLayout = new FlowLayout(FlowLayout.CENTER);
		JPanel southPanel = new JPanel(southLayout);
		JButton confirmButton = new JButton("确定");
		confirmButton.setActionCommand("confirm");
		confirmButton.addActionListener(this);
		southPanel.add(confirmButton);
		this.add(southPanel, BorderLayout.SOUTH);
	}

	/**
	 * 命令执行(ActionListener)接口实现
	 */
	public void actionPerformed(ActionEvent e) {
		// 获取按钮或菜单命令
		String actionCommand = e.getActionCommand();
		if (actionCommand.equals("confirm")) {
			oldExcelFilePath = oldExcelFileInput.getText();
			newExcelFilePath = newExcelFileInput.getText();
			if (oldExcelFilePath == null || 
				oldExcelFilePath.equals("")) {
				JOptionPane.showMessageDialog(this, "请选择旧BOM的EXCEL文件", "提示", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			if (newExcelFilePath == null || 
				newExcelFilePath.equals("")) {
				JOptionPane.showMessageDialog(this, "请选择旧BOM的EXCEL文件", "提示", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			
	    	List<String[]> resultOld = null;
			try {
				resultOld = ExcelRead.readExcel(oldExcelFilePath);
		    	System.out.println(ExcelRead.formatJson(JSON.toJSONString(resultOld)));
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
	    	
	    	System.out.println("--------------------------------");
	    	
	    	List<String[]> resultNew = null;
			try {
				resultNew = ExcelRead.readExcel(newExcelFilePath);
		    	System.out.println(ExcelRead.formatJson(JSON.toJSONString(resultNew)));
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			try {
				String oldPath = newExcelFilePath.substring(0, newExcelFilePath.lastIndexOf("\\") + 1);  
		           
				CreateWorkbookInPOI.CreateDiffResultExcel(oldPath, resultOld, resultNew);
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			this.dispose();
		} else if (actionCommand.equals("oldFolderSelect")) {
			// 用JFileChooser实现选择文件夹
			String initFolderPath = oldExcelFilePath;
			if (initFolderPath == null || 
				initFolderPath.equals("")) {
				initFolderPath = newExcelFilePath;
			}
			JFileChooser fileChooser = new JFileChooser(initFolderPath);
			fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
			fileChooser.showDialog(this, "选择旧BOM的EXCEL文件");
			try {
				File file = fileChooser.getSelectedFile();
				if (file != null) {
					oldExcelFilePath = file.getCanonicalPath();
					oldExcelFileInput.setText(oldExcelFilePath);
				}
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		} else if (actionCommand.equals("newFolderSelect")) {
			// 用JFileChooser实现选择文件夹
			String initFolderPath = newExcelFilePath;
			if (initFolderPath == null || initFolderPath.equals("")) {
				initFolderPath = oldExcelFilePath;
			}
			JFileChooser fileChooser = new JFileChooser(initFolderPath);
			fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
			fileChooser.showDialog(this, "选择新BOM的EXCEL文件");
			try {
				File file = fileChooser.getSelectedFile();
				if (file != null) {
					newExcelFilePath = file.getCanonicalPath();
					newExcelFileInput.setText(newExcelFilePath);
				}
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		} else {
			System.out.println(actionCommand);
		}
	}

	/**
	 * 返回旧的文件夹路径
	 * 
	 * @return 旧的文件夹路径
	 */
	public String getOldFolderPath() {
		return oldExcelFilePath;
	}

	/**
	 * 返回新的文件夹路径
	 * 
	 * @return 新的文件夹路径
	 */
	public String getNewFolderPath() {
		return newExcelFilePath;
	}

	/**
	 * 初始设定已经选择的路径
	 * 
	 * @param oldFolderPath
	 */
	public void setOldFolderPath(String oldFolderPath) {
		this.oldExcelFilePath = oldFolderPath;
		this.oldExcelFileInput.setText(oldFolderPath);
	}

	/**
	 * 初始设定已经选择的路径
	 * 
	 * @param newFolderPath
	 */
	public void setNewFolderPath(String newFolderPath) {
		this.newExcelFilePath = newFolderPath;
		this.newExcelFileInput.setText(newFolderPath);
	}

}
