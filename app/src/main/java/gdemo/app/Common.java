/**
 * 
 */
package gdemo.app;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.Vector;

/**
 * @author hammmi
 *
 */

public final class Common {

	public static String logo = "######################################################################\n" + 
			"#####                                                            #####\n" +
			"#####                        Joiner Leaver                       #####\n" +
			"#####                                                            #####\n" +
			"######################################################################\n";
	
	public static File sourcePath = new File("C:\\Users\\hammmi\\Documents\\CGI Betriebsrat\\JoinerLeaver");    
	public static File targetPath = new File("C:\\Users\\hammmi\\Documents\\CGI Betriebsrat\\JoinerLeaverXX");  
	public static String postfix = ".xlsm";
	
	static List<File> listDir(File dir) {
		
		if(dir == null) {
			dir = sourcePath;
		}
		List<File> sourceFiles = Arrays.asList(dir.listFiles());
		List<File> targetFiles = new Vector<File>();

		if (!sourceFiles.isEmpty()) {
			for (File file : sourceFiles) {
				if (file.getName().endsWith(Common.postfix) && file.exists() && file.isFile()) {
					if (file.getName().endsWith(Common.postfix)) {
						targetFiles.add(file);
					}
				}

			}
		}
		return targetFiles;
	}
}
