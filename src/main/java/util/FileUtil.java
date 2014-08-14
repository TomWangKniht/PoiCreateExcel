package util;

public class FileUtil {

	public static String getSubFileName(String fullFileName) {
		int index = fullFileName.lastIndexOf(".");
		return index==-1?"":fullFileName.substring(index);
	}

}
