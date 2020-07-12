package com.windows.printer;

import java.awt.print.Book;
import java.awt.print.PageFormat;
import java.awt.print.Paper;
import java.awt.print.PrinterJob;
import java.io.*;
import java.util.Arrays;
import java.util.List;
import javax.print.*;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Copies;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.apache.pdfbox.multipdf.Splitter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.printing.PDFPrintable;
import org.apache.pdfbox.printing.Scaling;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

/**
 *  @date 2020/7/1 22:07
 */
@Controller
@RequestMapping("/print")
public class PrintController {
    public final int WIDTH = 595;

    public final int HEIGHT = 842;

    @SuppressWarnings("AlibabaMethodTooLong")
    @ResponseBody
    @RequestMapping(value = "printByPrinterName", method = RequestMethod.GET)
    public String print(HttpServletRequest request, HttpServletResponse response) {
        try {
            //文件所在文件夹
            String filePath = "D:\\printDirectory";
            String print = request.getParameter("printer");
            if (print == null || "".equals(print)) {
                throw new RuntimeException("参数为空");
            }

            PrintService[] services = PrinterJob.lookupPrintServices();
            PrintService service = null;
            for (int i = 0; i < services.length; i++) {
                if (print.equals(services[i].getName())) {
                    service = services[i];
                    break;
                }
            }
            if (service == null) {
                throw new RuntimeException("打印机不存在");
            }
            File directory = new File(filePath);
            if (!directory.exists()) {
                throw new RuntimeException("文件夹不存在");
            }
            if (directory.isFile()) {
                throw new RuntimeException("文件夹格式不正确");
            }
            String[] list = directory.list();
            if(list == null) {
                throw new RuntimeException("文件夹为空");
            }
            List<String> fileList = Arrays.asList(list);
            System.out.println("fileList: \n" + fileList.toString());
            for (String fileName : fileList) {
                String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
                System.out.println("后缀(suffix) : " + suffix);
                String file = filePath + File.separator + fileName;
                if("pdf".equals(suffix)) {
                    printPdf(service, file);
                } else if("doc".equals(suffix) || "docx".equals(suffix)) {
                    printDoc(service, file);
                } else if("jpg".equals(suffix) || "jpeg".equals(suffix)) {
                    printJpg(service, file);
                } else if("xls".equals(suffix) || "xlsx".equals(suffix)) {
                    printExcel(service, file);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            return "false";
        }
        return "success";
    }

    private void printPdf(PrintService service , String file) throws  Exception{
        InputStream in = new FileInputStream(file);
        PDDocument document = PDDocument.load(in);
        Splitter splitter = new Splitter();
        List<PDDocument> pages = splitter.split(document);
        //指定打印机
        PrinterJob job = PrinterJob.getPrinterJob();
        job.setPrintService(service);
        //纸张等设置
        Paper paper = new Paper();
        paper.setSize(WIDTH, HEIGHT);
        paper.setImageableArea(0, 0, WIDTH, HEIGHT);
        PageFormat pageFormat = new PageFormat();
        pageFormat.setPaper(paper);
        PDFPrintable pdfPrintable = new PDFPrintable(document, Scaling.ACTUAL_SIZE);
        Book book = new Book();
        book.append(pdfPrintable, pageFormat, pages.size());
        job.setPageable(book);
        job.print();
        document.close();
        in.close();
    }

    /**
     * jpg文件打印
     */
    private void printJpg(PrintService service, String file) throws Exception {
        InputStream fis = new FileInputStream(file);
        // 设置打印格式，如果未确定类型，可选择autosense
        DocFlavor flavor = DocFlavor.INPUT_STREAM.JPEG;
        // 设置打印参数
        PrintRequestAttributeSet aset = new HashPrintRequestAttributeSet();
        aset.add(new Copies(1));
        Doc doc = new SimpleDoc(fis, flavor, null);
        // 创建打印作业
        DocPrintJob job = service.createPrintJob();
        job.print(doc, aset);
        fis.close();
    }

    private void printDoc(PrintService service , String file) throws Exception {
        //初始化线程
        ComThread.InitSTA();
        ActiveXComponent word = new ActiveXComponent("Word.Application");
        //设置打印机名称
        word.setProperty("ActivePrinter", new Variant(service.getName()));
        // 这里Visible是控制文档打开后是可见还是不可见，若是静默打印，那么第三个参数就设为false就好了
        Dispatch.put(word, "Visible", new Variant(false));
        // 获取文档属性
        Dispatch document = word.getProperty("Documents").toDispatch();
        // 打开激活文挡
        Dispatch doc = Dispatch.call(document, "Open", file).toDispatch();

        Dispatch.callN(doc, "PrintOut");
        System.out.println("打印成功！");

        //word文档关闭
        Dispatch.call(doc, "Close", new Variant(0));
        //退出
        word.invoke("Quit", new Variant[0]);
        //释放资源
        ComThread.Release();
        ComThread.quitMainSTA();
    }

    private void printExcel(PrintService service, String file) {
        ComThread.InitSTA();
        ActiveXComponent xl = new ActiveXComponent("Excel.Application");
        // 不打开文档
        Dispatch.put(xl, "Visible", new Variant(false));
        Dispatch workbooks = xl.getProperty("Workbooks").toDispatch();
        Object[] object = new Object[8];
        object[0] = Variant.VT_MISSING;
        object[1] = Variant.VT_MISSING;
        object[2] = Variant.VT_MISSING;
        object[3] = new Boolean(false);
        object[4] = service.getName();
        object[5] = new Boolean(false);
        object[6] = Variant.VT_MISSING;
        object[7] = Variant.VT_MISSING;

        // 打开文档
        Dispatch excel = Dispatch.call(workbooks, "Open", file).toDispatch();
        // 每张表都横向打印2013-10-31
        Dispatch sheets = Dispatch.get((Dispatch) excel, "Sheets").toDispatch();
        // 获得几个sheet
        int count = Dispatch.get(sheets, "Count").getInt();
        for (int j = 1; j <= count; j++) {
            Dispatch sheet = Dispatch
                    .invoke(sheets, "Item", Dispatch.Get, new Object[]{new Integer(j)}, new int[1])
                    .toDispatch();
            Dispatch pageSetup = Dispatch.get(sheet, "PageSetup").toDispatch();
            Dispatch.put(pageSetup, "Orientation", new Variant(2));
            Dispatch.call(sheet, "PrintOut", object);
        }
        // 增加以下三行代码解决文件无法删除bug
        Dispatch.call(excel, "save");
        Dispatch.call(excel, "Close", new Variant(true));
        xl.invoke("Quit", new Variant[]{});

        // 始终释放资源
        ComThread.Release();
    }

}

