package com.wiwj.cbs;

import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

public class BOMDiff {

	/**
	 * test
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		final CompareExcelFileSelectDialog test = new CompareExcelFileSelectDialog(null);
		test.addWindowListener(new WindowAdapter() {
			public void windowClosing(WindowEvent e) {
				System.out.println(test.getOldFolderPath());
				System.out.println(test.getNewFolderPath());
				System.exit(0);
			}
		});
		test.setVisible(true);
	}

}
