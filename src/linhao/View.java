package linhao;


import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

public class View {

	static String filename,dirname;
	static JButton button1, button2, button3, button4;
	static JTextField text1, text2;
	static JFileChooser jfc1,jfc2;
	static Model m;

	public static void createView() {//==================================================================界面呈现

		JLabel label1 = new JLabel("汇总文件:");
		JLabel label2 = new JLabel("输出路径:");
		JLabel label3 = new JLabel("");
		text1 = new JTextField();
		text2 = new JTextField();
		button1 = new JButton("打开");
		button2 = new JButton("选择");
		button3 = new JButton("一键生成");
		button4 = new JButton("清除");
		m = new Model();

		button1.addActionListener(new ActionListener() {//======================================================按钮：汇总文件
			@Override
			public void actionPerformed(ActionEvent e) {
				jfc1 = new JFileChooser();
				jfc1.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc1.setCurrentDirectory(new File("."));
				jfc1.showDialog(new JLabel(), "选择");
				text1.setText(jfc1.getSelectedFile().getAbsolutePath());
				filename = jfc1.getSelectedFile().getAbsolutePath();
			}
		});

		button2.addActionListener(new ActionListener() {//======================================================按钮：输出路径
			@Override
			public void actionPerformed(ActionEvent e) {
				jfc2 = new JFileChooser();
				jfc2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				jfc2.setCurrentDirectory(new File("."));
				jfc2.showDialog(new JLabel(), "选择");
				text2.setText(jfc2.getSelectedFile().getAbsolutePath());
				dirname = jfc2.getSelectedFile().getAbsolutePath();
			}
		});

		button3.addActionListener(new ActionListener() {//======================================================按钮：开始转换
			@Override
			public void actionPerformed(ActionEvent e) {
				if (filename == null) {
					JOptionPane.showMessageDialog(null, "请先选择汇总文件");
				}else	if (dirname == null) {
					JOptionPane.showMessageDialog(null, "请先选择输出目录");
				}else	if (filename != null && dirname != null) {
					try {
						Model m = new Model();
						m.readExcel(filename,dirname);
						File dir2 = new File(dirname);
						Runtime.getRuntime().exec("cmd /c start explorer " + dir2);
						System.exit(0);
					} catch (Exception e1) {
						e1.printStackTrace();
					}
				}
			}
		});
		
		button4.addActionListener(new ActionListener() {//======================================================按钮：清空选择
			@Override
			public void actionPerformed(ActionEvent e) {
				text1.setText("");
				text2.setText("");
				filename = null;
				dirname = null;
			}
		});

		JPanel panel = new JPanel();
		panel.setLayout(new GridLayout(3, 3));
		panel.add(label1);
		panel.add(text1);
		panel.add(button1);
		panel.add(label2);
		panel.add(text2);
		panel.add(button2);
		panel.add(label3);
		panel.add(button3);
		panel.add(button4);

		JFrame frame = new JFrame("营业部信贷室v1.0");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(300, 120);
		frame.setContentPane(panel);
		frame.setVisible(true);

	}

	public static void main(String[] args) {
		// 单独线程管理UI
		createView();
	}

}
