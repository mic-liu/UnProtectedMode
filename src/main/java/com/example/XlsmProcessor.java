package com.example;

import java.io.*;
import java.nio.charset.Charset;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class XlsmProcessor {
    public static void main(String[] args) {
        if (args.length != 1) {
            System.out.println("请将Excel文件拖拽到本程序上！");
            return;
        }

        String inputFile = args[0].toLowerCase();
        if (!inputFile.endsWith(".xlsm") && !inputFile.endsWith(".xlsx") && !inputFile.endsWith(".xls")) {
            System.out.println("请拖拽Excel文件（.xlsm、.xlsx或.xls）！");
            return;
        }

        // 构造输出文件名
        String outputFile = args[0].substring(0, args[0].lastIndexOf(".")) + "_unprotected" + args[0].substring(args[0].lastIndexOf("."));

        try {
            // 创建临时目录
            File tempDir = new File("temp");
            tempDir.mkdir();

            // 解压Excel文件
            unzipFile(args[0], tempDir.getPath());

            // 处理所有XML文件
            processXmlFiles(tempDir);

            // 重新打包
            zipDirectory(tempDir, outputFile);

            // 清理临时文件
            deleteDirectory(tempDir);

            System.out.println("处理完成！输出文件：" + outputFile);
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("处理失败！");
        }
    }

    private static void unzipFile(String zipFile, String outputFolder) throws IOException {
        byte[] buffer = new byte[1024];
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFile))) {
            ZipEntry zipEntry = zis.getNextEntry();
            while (zipEntry != null) {
                String fileName = zipEntry.getName();
                File newFile = new File(outputFolder + File.separator + fileName);
                
                // 创建目录
                new File(newFile.getParent()).mkdirs();

                if (!zipEntry.isDirectory()) {
                    try (FileOutputStream fos = new FileOutputStream(newFile)) {
                        int len;
                        while ((len = zis.read(buffer)) > 0) {
                            fos.write(buffer, 0, len);
                        }
                    }
                }
                zipEntry = zis.getNextEntry();
            }
        }
    }

    private static void processXmlFiles(File directory) throws IOException {
        // 添加workbookProtection的正则表达式
        Pattern sheetProtectionPattern = Pattern.compile("<sheetProtection\\b[^>]*\\/>");
        Pattern workbookProtectionPattern = Pattern.compile("<workbookProtection\\b[^>]*\\/>");
        
        if (directory.isDirectory()) {
            for (File file : directory.listFiles()) {
                if (file.isDirectory()) {
                    processXmlFiles(file);
                } else if (file.getName().endsWith(".xml")) {
                    // 读取文件内容
                    String content = readFile(file);
                    // 替换匹配的内容
                    content = sheetProtectionPattern.matcher(content).replaceAll("");
                    content = workbookProtectionPattern.matcher(content).replaceAll("");
                    // 写回文件
                    writeFile(file, content);
                }
            }
        }
    }

    private static String readFile(File file) throws IOException {
        String encoding = "UTF-8";

        try (FileInputStream fis = new FileInputStream(file);
             BufferedInputStream bis = new BufferedInputStream(fis)) {
            
            bis.mark(4);
            byte[] bom = new byte[4];
            int n = bis.read(bom, 0, 4);
            
            if (n >= 3 && bom[0] == (byte)0xEF && bom[1] == (byte)0xBB && bom[2] == (byte)0xBF) {
                encoding = "UTF-8";
                bis.reset();
                bis.skip(3);
            } else if (n >= 2 && bom[0] == (byte)0xFF && bom[1] == (byte)0xFE) {
                encoding = "UTF-16LE";
                bis.reset();
                bis.skip(2);
            } else if (n >= 2 && bom[0] == (byte)0xFE && bom[1] == (byte)0xFF) {
                encoding = "UTF-16BE";
                bis.reset();
                bis.skip(2);
            } else {
                bis.reset();
            }
            
            StringBuilder content = new StringBuilder();
            char[] buffer = new char[8192];
            try (BufferedReader reader = new BufferedReader(new InputStreamReader(bis, encoding))) {
                int len;
                while ((len = reader.read(buffer)) != -1) {
                    content.append(buffer, 0, len);
                }
            }
            return content.toString();
        }
    }

    private static void writeFile(File file, String content) throws IOException {
        // First read the file to detect original encoding
        String encoding = "UTF-8";
        byte[] bom = null;

        try (FileInputStream fis = new FileInputStream(file);
             BufferedInputStream bis = new BufferedInputStream(fis)) {
            
            bis.mark(4);
            byte[] bomCheck = new byte[4];
            int n = bis.read(bomCheck, 0, 4);
            
            if (n >= 3 && bomCheck[0] == (byte)0xEF && bomCheck[1] == (byte)0xBB && bomCheck[2] == (byte)0xBF) {
                encoding = "UTF-8";
                bom = new byte[]{(byte)0xEF, (byte)0xBB, (byte)0xBF};
            } else if (n >= 2 && bomCheck[0] == (byte)0xFF && bomCheck[1] == (byte)0xFE) {
                encoding = "UTF-16LE";
                bom = new byte[]{(byte)0xFF, (byte)0xFE};
            } else if (n >= 2 && bomCheck[0] == (byte)0xFE && bomCheck[1] == (byte)0xFF) {
                encoding = "UTF-16BE";
                bom = new byte[]{(byte)0xFE, (byte)0xFF};
            }
        }

        try (FileOutputStream fos = new FileOutputStream(file)) {
            // Write BOM if it existed in the original file
            if (bom != null) {
                fos.write(bom);
            }
            
            // Write content with detected encoding
            try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(fos, encoding))) {
                writer.write(content);
                writer.flush();
            }
        }
    }

    private static void zipDirectory(File directory, String zipFile) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(zipFile);
             ZipOutputStream zos = new ZipOutputStream(fos)) {
            zipFile(directory, "", zos);
        }
    }

    private static void zipFile(File directory, String baseName, ZipOutputStream zos) throws IOException {
        File[] files = directory.listFiles();
        byte[] buffer = new byte[1024];
        
        for (File file : files) {
            String entryName = baseName + (baseName.isEmpty() ? "" : File.separator) + file.getName();
            
            if (file.isDirectory()) {
                zipFile(file, entryName, zos);
                continue;
            }
            
            ZipEntry ze = new ZipEntry(entryName);
            zos.putNextEntry(ze);
            
            try (FileInputStream in = new FileInputStream(file)) {
                int len;
                while ((len = in.read(buffer)) > 0) {
                    zos.write(buffer, 0, len);
                }
            }
        }
    }

    private static void deleteDirectory(File directory) {
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    deleteDirectory(file);
                } else {
                    file.delete();
                }
            }
        }
        directory.delete();
    }
}