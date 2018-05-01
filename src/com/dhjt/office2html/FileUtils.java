package com.dhjt.office2html;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;

/**
 * File工具类，可以使用apache或者其他的工具类实现该类的作用
 * @author DHJT 2018年5月1日 下午12:38:23
 *
 */
public class FileUtils {
	private static final String ENCODING = "GB2312";// UTF-8

	public static String GetFileExt(String path) {
		String ext = null;
		int i = path.lastIndexOf('.');
		if (i > 0 && i < path.length() - 1) {
			ext = path.substring(i + 1).toLowerCase();
		}
		return ext;
	}

	public static void createHtmlDir(String path) {
		int i = path.lastIndexOf('/');
		String dirPath = "";
		if (i > 0 && i < path.length() - 1) {
			dirPath = path.substring(0, i).toLowerCase();
		}
		File dir = new File(dirPath);
		if (!dir.exists()) {
			dir.mkdirs();
		}

	}

	public static void createDir(String dirPath) {
		File dir = new File(dirPath);
		if (!dir.exists()) {
			dir.mkdirs();
		}

	}

	public static void writeFile(String content, String path) {
		createHtmlDir(path);
		OutputStream os = null;
		BufferedWriter bw = null;
		try {
			File file = new File(path);
			if (!file.exists()) {
				file.createNewFile();
			}
			os = new FileOutputStream(file);
			bw = new BufferedWriter(new OutputStreamWriter(os, ENCODING));
			bw.write(content);
		} catch (FileNotFoundException fnfe) {
			fnfe.printStackTrace();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		} finally {
			try {
				if (bw != null)
					bw.close();
				if (os != null)
					os.close();
			} catch (IOException ie) {
				ie.printStackTrace();
			}
		}
	}

}
