package modifyExcel;

import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.util.Random;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class fmMain {

	private JFrame jfMain=null;
	
	private JLabel jlWeibo=new JLabel("微博：");
	private JLabel jlSWeibo=new JLabel("原微博：");
	
	private JLabel jlTopic=new JLabel("主题：");
	private JLabel jlWord=new JLabel("关键词：");
	private JLabel jlLevel=new JLabel("自杀程度：");
	private JLabel jlSign=new JLabel("自杀讯号：");
	
	private JButton jbNo=new JButton("无");
	private JButton jbRandom1=new JButton("随机");
	private JButton jbRandom2=new JButton("随机");
	private JButton jbRandom3=new JButton("随机");
	private JButton jbPrevious=new JButton("上条");
	private JButton jbNext=new JButton("下条");
	private JButton jbSaveNext=new JButton("保存并下条");
	
	private JTextField jtfTopic=new JTextField();
	private JTextField jtfWord=new JTextField();
	private JTextField jtfLevel=new JTextField();
	private JTextField jtfSign=new JTextField();
	
	private JTextArea jtaWeibo=new JTextArea();
	private JTextArea jtaSWeibo=new JTextArea();
	
	private JScrollPane jspWeibo=new JScrollPane(jtaWeibo,JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
	private JScrollPane jspSWeibo=new JScrollPane(jtaSWeibo,JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
	
	private File f;
	private WritableWorkbook copy;
	private WritableSheet sheet=null;
	private int cline=1;	//行
	private int start=1;
	private int end=652;
	
	public fmMain() {
		jfMain=new JFrame("评价自杀微博");
		
		Font fnt = new Font("微软雅黑",Font.PLAIN +Font.BOLD,14);
		jtaWeibo.setFont(fnt);
		jtaSWeibo.setFont(fnt);
		jtfTopic.setFont(fnt);
		jtfWord.setFont(fnt);
		
		jtaWeibo.setEditable(false);
		jtaSWeibo.setEditable(false);
		jtaWeibo.setLineWrap(true);//设置自动换行
		jtaWeibo.setWrapStyleWord(true);
		jtaSWeibo.setLineWrap(true);//设置自动换行
		jtaSWeibo.setWrapStyleWord(true);
		
		jlWeibo.setBounds(20, 15, 60, 20);
		jspWeibo.setBounds(20, 40, 255, 80);
		jlSWeibo.setBounds(20, 125, 60, 20);
		jspSWeibo.setBounds(20, 150, 255, 80);
		jlTopic.setBounds(20, 240, 40, 20);
		jtfTopic.setBounds(65, 240, 145, 20);
		jbRandom1.setBounds(215, 240, 60, 20);
		jlWord.setBounds(20, 270, 55, 20);
		jtfWord.setBounds(80, 270, 140, 20);
		jbNo.setBounds(225, 270, 50, 20);
		jlLevel.setBounds(20, 300, 70, 20);
		jtfLevel.setBounds(95, 300, 115, 20);
		jbRandom2.setBounds(215, 300, 60, 20);
		jlSign.setBounds(20, 330, 70, 20);
		jtfSign.setBounds(95, 330, 115, 20);
		jbRandom3.setBounds(215, 330, 60, 20);
		jbPrevious.setBounds(20, 365, 60, 30);
		jbNext.setBounds(90, 365, 60, 30);
		jbSaveNext.setBounds(165, 365, 110, 30);
		
		jfMain.setLayout(null);
		
		jfMain.add(jlWeibo);
		jfMain.add(jspWeibo);
		jfMain.add(jlSWeibo);
		jfMain.add(jspSWeibo);
		jfMain.add(jlTopic);
		jfMain.add(jtfTopic);
		jfMain.add(jbRandom1);
		jfMain.add(jlWord);
		jfMain.add(jtfWord);
		jfMain.add(jbNo);
		jfMain.add(jlLevel);
		jfMain.add(jtfLevel);
		jfMain.add(jbRandom2);
		jfMain.add(jlSign);
		jfMain.add(jtfSign);
		jfMain.add(jbRandom3);
		jfMain.add(jbPrevious);
		jfMain.add(jbNext);
		jfMain.add(jbSaveNext);
		
		jfMain.setResizable(false);
		jfMain.setSize(300, 440);
		jfMain.setLocationRelativeTo(null);
		jfMain.setVisible(true);
		
		f=new File("D:"+File.separator+"output.xls");
		if(!f.exists()){
			f=new File("D:"+File.separator+"homework.xls");
		}
		
		//operate excel
		try {
			Workbook book=Workbook.getWorkbook(f);
			copy = Workbook.createWorkbook(new File("D:"+File.separator+"output.xls"), book);
			book.close();
			sheet=copy.getSheet(0);
			jtaWeibo.setText(sheet.getCell(3, cline).getContents());
			jtaSWeibo.setText(sheet.getCell(4, cline).getContents());
			jtfTopic.setText(sheet.getCell(5, cline).getContents());
			jtfWord.setText(sheet.getCell(6, cline).getContents());
			jtfLevel.setText(sheet.getCell(7, cline).getContents());
			jtfSign.setText(sheet.getCell(8, cline).getContents());
			
			jfMain.setTitle("第 "+cline+" 条微博");
		} catch (Exception e1) {
			JOptionPane.showMessageDialog(jfMain.getContentPane(),
				       "打开文件错误!", "注意！", JOptionPane.WARNING_MESSAGE);
			System.exit(1);
		}
		
		jfMain.addWindowListener(
				new WindowAdapter(){
					/**
					 * it happens when close the windows
					 */
					public void windowClosing(WindowEvent e){
						try {
							copy.write();
							copy.close();
						} catch (Exception e1) {
							e1.printStackTrace();
						}
						System.exit(1) ;
					}
				}	
				);
		
		jbPrevious.addActionListener(
				new ActionListener(){
					public void actionPerformed(ActionEvent e) {
						if(e.getSource()==jbPrevious){
							if(cline==start){
								JOptionPane.showMessageDialog(jfMain.getContentPane(),
									       "已到最前!", "注意！", JOptionPane.WARNING_MESSAGE);
							}else{
								cline-=1;
								jtaWeibo.setText(sheet.getCell(3, cline).getContents());
								jtaSWeibo.setText(sheet.getCell(4, cline).getContents());
								jtfTopic.setText(sheet.getCell(5, cline).getContents());
								jtfWord.setText(sheet.getCell(6, cline).getContents());
								jtfLevel.setText(sheet.getCell(7, cline).getContents());
								jtfSign.setText(sheet.getCell(8, cline).getContents());
								
								jfMain.setTitle("第 "+cline+" 条微博");
							}
							
						}	
					}
				}
				);
		
		jbNext.addActionListener(
				new ActionListener(){
					public void actionPerformed(ActionEvent e) {
						if(e.getSource()==jbNext){
							if(cline==end){
								JOptionPane.showMessageDialog(jfMain.getContentPane(),
									       "已到最后!", "注意！", JOptionPane.WARNING_MESSAGE);
							}else{
								cline+=1;
								jtaWeibo.setText(sheet.getCell(3, cline).getContents());
								jtaSWeibo.setText(sheet.getCell(4, cline).getContents());
								jtfTopic.setText(sheet.getCell(5, cline).getContents());
								jtfWord.setText(sheet.getCell(6, cline).getContents());
								jtfLevel.setText(sheet.getCell(7, cline).getContents());
								jtfSign.setText(sheet.getCell(8, cline).getContents());
								
								jfMain.setTitle("第 "+cline+" 条微博");
							}
							
						}	
					}
				}
				);
		
		jbSaveNext.addActionListener(
				new ActionListener(){
					public void actionPerformed(ActionEvent e) {
						if(e.getSource()==jbSaveNext){
							if(cline==end){
								JOptionPane.showMessageDialog(jfMain.getContentPane(),
									       "已到最后!", "注意！", JOptionPane.WARNING_MESSAGE);
							}else{
								Label l;
								l=new Label(5, cline,jtfTopic.getText());
								try {
									sheet.addCell(l);
								} catch (Exception e1) {						
									JOptionPane.showMessageDialog(jfMain.getContentPane(),
										       "修改失败!", "注意！", JOptionPane.WARNING_MESSAGE);
								}
								l=new Label(6, cline,jtfWord.getText());
								try {
									sheet.addCell(l);
								} catch (Exception e1) {						
									JOptionPane.showMessageDialog(jfMain.getContentPane(),
										       "修改失败!", "注意！", JOptionPane.WARNING_MESSAGE);
								}
								l=new Label(7, cline,jtfLevel.getText());
								try {
									sheet.addCell(l);
								} catch (Exception e1) {						
									JOptionPane.showMessageDialog(jfMain.getContentPane(),
										       "修改失败!", "注意！", JOptionPane.WARNING_MESSAGE);
								}
								l=new Label(8, cline,jtfSign.getText());
								try {
									sheet.addCell(l);
								} catch (Exception e1) {						
									JOptionPane.showMessageDialog(jfMain.getContentPane(),
										       "修改失败!", "注意！", JOptionPane.WARNING_MESSAGE);
								}
								
								cline+=1;
								jtaWeibo.setText(sheet.getCell(3, cline).getContents());
								jtaSWeibo.setText(sheet.getCell(4, cline).getContents());
								jtfTopic.setText(sheet.getCell(5, cline).getContents());
								jtfWord.setText(sheet.getCell(6, cline).getContents());
								jtfLevel.setText(sheet.getCell(7, cline).getContents());
								jtfSign.setText(sheet.getCell(8, cline).getContents());
								
								jfMain.setTitle("第 "+cline+" 条微博");
							}
							
						}	
					}
				}
				);
		
		jbNo.addActionListener(
				new ActionListener(){
					public void actionPerformed(ActionEvent e) {
						if(e.getSource()==jbNo){
							jtfWord.setText("999");
						}
					}
				}
				);
		
		jbRandom1.addActionListener(
				new ActionListener(){
					public void actionPerformed(ActionEvent e) {
						if(e.getSource()==jbRandom1){
							jtfTopic.setText("NA");
						}
					}
				}
				);
		
		jbRandom2.addActionListener(
				new ActionListener(){
					public void actionPerformed(ActionEvent e) {
						if(e.getSource()==jbRandom2){
							
//							Random r = new Random();
//							int i=r.nextInt(4);	//i 为[0,3]的随机数
							jtfLevel.setText("0");
						}
					}
				}
				);
		
		jbRandom3.addActionListener(
				new ActionListener(){
					public void actionPerformed(ActionEvent e) {
						if(e.getSource()==jbRandom3){
							jtfSign.setText("0");
						}
					}
				}
				);
		
	}
	
	public static void main(String args[]){
		new fmMain();
	}
}
