package com.itod;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.util.Units;
import org.apache.commons.io.FilenameUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class App {
    public static void main(String[] args) {
       // String folderPath = "/Users/soma/imgs";
      //  String outputPath = "/Users/soma/imgs/doc/output.docx";

        if (args.length < 2) {
            System.out.println("Usage: java FolderImagesToWord <folderPath> <outputFilePath>");
            return;
        }

        String folderPath = args[0];
        String outputPath = args[1];

        try (XWPFDocument doc = new XWPFDocument()) {
            File folder = new File(folderPath);
            File[] listOfFiles = folder.listFiles();

            if (listOfFiles != null) {
                for (File file : listOfFiles) {
                    if (file.isFile() && isImageFile(file)) {
                        String title = getTitleFromFilename(file.getName());
                        addTitleToDocument(doc, title);
                        XWPFParagraph paragraph = doc.createParagraph();
                        XWPFRun run = paragraph.createRun();
                        run.addBreak();
                        try (FileInputStream is = new FileInputStream(file)) {
                            run.addPicture(is, getPictureType(file), file.getName(), Units.toEMU(500), Units.toEMU(500)); // Adjust image size as needed
                        }
                        addPageBreak(doc);
                    }
                }
            }

            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                doc.write(out);
                System.out.println("Word document created successfully.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void addTitleToDocument(XWPFDocument document, String title) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(title);
        run.setBold(true);
        run.setFontSize(14);
    }


    private static String getTitleFromFilename(String filename) {
        if (filename == null || filename.isEmpty()) {
            return "";
        }

        // Remove file extension
        int dotIndex = filename.lastIndexOf('.');
        if (dotIndex > 0) {
            filename = filename.substring(0, dotIndex);
        }

        // Remove "Capture" from the filename
        filename = filename.replace("Capture", "").trim();

        return filename;
    }

    private static void addPageBreak(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.addBreak(org.apache.poi.xwpf.usermodel.BreakType.PAGE);
    }

    private static boolean isImageFile(File file) {
        String extension = FilenameUtils.getExtension(file.getName()).toLowerCase();
        
    
        return extension.equals("jpg") || extension.equals("jpeg") || extension.equals("gif") || extension.equals("bmp") ||  extension.equals("png") ;
    }

    private static int getPictureType(File file) {
        String extension = FilenameUtils.getExtension(file.getName()).toLowerCase();
        switch (extension) {
            case "emf": return XWPFDocument.PICTURE_TYPE_EMF;
            case "wmf": return XWPFDocument.PICTURE_TYPE_WMF;
            case "pict": return XWPFDocument.PICTURE_TYPE_PICT;
            case "jpeg":
            case "jpg": return XWPFDocument.PICTURE_TYPE_JPEG;
            case "png": return XWPFDocument.PICTURE_TYPE_PNG;
            case "dib": return XWPFDocument.PICTURE_TYPE_DIB;
            case "gif": return XWPFDocument.PICTURE_TYPE_GIF;
            case "tiff": return XWPFDocument.PICTURE_TYPE_TIFF;
            case "eps": return XWPFDocument.PICTURE_TYPE_EPS;
            case "bmp": return XWPFDocument.PICTURE_TYPE_BMP;
            case "wpg": return XWPFDocument.PICTURE_TYPE_WPG;
            default: return XWPFDocument.PICTURE_TYPE_JPEG;
        }
    }
}
