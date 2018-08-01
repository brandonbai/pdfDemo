package com.github.brandonbai.pdfDemo.util;

import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.PDFMergerUtility;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Arrays;
import java.util.Comparator;

/**
 * pdfbox 工具类
 * @author brandon
 * @since 2017-08-01
 */
public class PdfBoxUtil {

    /**
     * merge
     * @throws IOException
     */
    public static void merge(String folderName, String destPath) throws IOException {
        PDFMergerUtility mergePdf = new PDFMergerUtility();
        String[] filesInFolder = getFiles(folderName);
        Arrays.sort(filesInFolder, new Comparator<String>() {
            public int compare(String o1, String o2) {
                return o1.compareTo(o2);
            }
        });
        for (int i = 0; i < filesInFolder.length; i++) {
            mergePdf.addSource(folderName + File.separator + filesInFolder[i    ]);
        }
        mergePdf.setDestinationFileName(destPath);
        mergePdf.mergeDocuments(MemoryUsageSetting.setupMainMemoryOnly());
    }

    /**
     *
     * @param folder
     * @return
     * @throws IOException
     */
    private static String[] getFiles(String folder) throws IOException {
        File _folder = new File(folder);
        String[] filesInFolder;

        if (_folder.isDirectory()) {

            filesInFolder = _folder.list(new FilenameFilter() {

                public boolean accept(File dir, String name) {
                    return name.endsWith(".pdf");
                }

            });
            return filesInFolder;
        } else {
            throw new IOException("Path is not a directory");
        }
    }

}
