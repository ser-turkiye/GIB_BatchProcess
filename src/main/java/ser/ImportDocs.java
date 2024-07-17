//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by FernFlower decompiler)
//

package ser;

import com.microsoft.schemas.office.visio.x2012.main.DocumentSettingsType;
import com.ser.blueline.*;
import com.ser.blueline.metaDataComponents.*;
import com.ser.blueline.modifiablemetadata.IArchiveFolderClassModifiable;
import com.ser.foldermanager.IFolder;
import com.ser.foldermanager.IFolderConnection;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.commons.io.output.FileWriterWithEncoding;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ImportDocs extends UnifiedAgent {
    Logger log = LogManager.getLogger(this.getClass().getName());
    String nameDescriptor1 = "ccmPrjDocNumber";
    String searchClassName = "Search Engineering Documents";
    ISession ses = null;
    IDocumentServer srv = null;
    String CIFNumber = "";
    String AccountNumber = "";
    String CustomerType = "";
    String FolderName = "";
    String FileName = "";
    String DocumentType = "";
    String DocumentDate = "";
    String StoredDate = "";
    String ProcessID;
    String LogFilePath = "c:/tmp2/bulk/";
    IStringMatrix settingsMatrix = null;
    List<String> importLog = new ArrayList<>();
    Integer cnt = 0;
    Integer importedCount = 0;
    Integer totalCount = 0;

    public ImportDocs() {
    }

    protected Object execute() {
        this.log.info("Initiate the agent");
        if (this.getEventDocument() == null) {
            return this.resultError("Null Document object.");
        } else {
            ses = getSes();
            srv = ses.getDocumentServer();
            settingsMatrix = getDocumentServer().getStringMatrix("GibFolderAllDocType", ses);
            ProcessID = UUID.randomUUID().toString();

            IDocument ldoc = this.getEventDocument();
            File logFile = new File(LogFilePath + "GIB_IMPORT_LOG_" + ProcessID + ".txt");

            HashMap<Integer, String> flds = new HashMap();
            flds.put(0, "CIF");
            flds.put(1, "AccountNumber");
            flds.put(2, "DocumentType");
            flds.put(3, "DocumentDate");
            flds.put(4, "StoredDate");
            flds.put(5, "CustomerType");
            flds.put(6, "FilePath");

            try {
                //addToLog(logFile,"Line1");
                String excelPath = this.exportDocumentContent(ldoc, "C:/tmp2/bulk/import");

                //String excelPath = "C:/tmp2/bulk/import/Migration04.xlsx";
                FileInputStream fist = new FileInputStream(excelPath);
                this.log.info("Exported excel file to path:" + excelPath);
                Workbook wrkb = new XSSFWorkbook(fist);
                HashMap<String, Row> list = listOfDocuments(wrkb);
                List<String> fields = fieldsOfDocuments(wrkb);
                fist.close();
                for(String field : fields){
                    flds.put(cnt, field);
                    cnt++;
                }
                totalCount = list.size() + 1;
                Iterator var8 = list.entrySet().iterator();
                String docDate, storedDate, docType, custType, filePath, rowNr;
                this.log.info("Start second loop.");
                while(var8.hasNext()) {
                    List<String> nodeNames = new ArrayList<>();
                    Map.Entry<String, Row> line = (Map.Entry)var8.next();
                    Row row = (Row)line.getValue();
                    rowNr = row.getCell(0).getStringCellValue().trim();
                    CIFNumber = row.getCell(1).getStringCellValue().trim();
                    log.info("Import Doc Info : " + CIFNumber);
                    AccountNumber = row.getCell(2).getStringCellValue().trim();
                    docType = row.getCell(3).getStringCellValue().trim();
                    docDate = row.getCell(4).getStringCellValue().trim();
                    storedDate = row.getCell(5).getStringCellValue().trim();
                    custType = row.getCell(6).getStringCellValue().trim();
                    filePath = row.getCell(7).getStringCellValue().trim();
                    CustomerType = (custType.equals("C") ? "corp" : "retail");
                    FileName = new File(filePath).getName();
                    if(CIFNumber.isEmpty()){
                        addToLog(logFile,"NOT IMPORTED (" + FileName + ") CIF NO IS EMPTY : " + CIFNumber);
                        continue;
                    }
                    if(docType.isEmpty()){
                        addToLog(logFile,"NOT IMPORTED (" + FileName + ") DOC TYPE IS EMPTY : " + docType);
                        continue;
                    }
                    if(filePath.isEmpty()){
                        addToLog(logFile,"NOT IMPORTED (" + FileName + ") FILE PATH IS EMPTY : " + filePath);
                        continue;
                    }
                    log.info("Import Doc Info : " + CIFNumber);
                    log.info("Import Doc Info : " + AccountNumber);
                    log.info("Import Doc Info : " + filePath);
                    log.info("Import Doc Info : " + FileName);

                    if(Objects.equals(FileName, "") || FileName.isEmpty()){
                        addToLog(logFile,"NOT IMPORTED (" + FileName + ") FILENAME IS EMPTY : " + FileName);
                        //throw new Exception("FileName is empty for filepath : " + filePath);
                    }

                    addToLog(logFile,"IMPORT STARTED FOR (" + FileName + ")  : " + CIFNumber + " / " + AccountNumber);

                    Date date1 = new SimpleDateFormat("dd.MM.yyyy").parse(docDate);
                    DateFormat dt1 = new SimpleDateFormat("yyyyMMdd");

                    Date date2 = new SimpleDateFormat("dd.MM.yyyy HH:mm").parse(storedDate);
                    SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMddhhmmssS");
                    String strDate = formatter.format(date2);

                    DocumentDate = dt1.format(date1);
                    StoredDate = strDate;
                    log.info("GIBDoc Import Doc Info stored date : " + StoredDate);

                    FolderName = getFolderNameFromGVList(docType,CustomerType);
                    if(Objects.equals(FolderName, "") || FolderName.isEmpty()){
                        addToLog(logFile,"> NOT IMPORTED (" + FileName + ") FOLDER NAME NOT FOUND : " + FolderName);
                        continue;
                        //throw new Exception("Folder not found for : " + docType);
                    }
                    DocumentType = getDocumentTypeFromGVList(docType,FolderName);

                    log.info("GIBDoc Import Doc found Class for Folder : " + FolderName);
                    String ClassID = getCategoryIDFromGVlist(FolderName);
                    if(Objects.equals(ClassID, "") || ClassID.isEmpty()){
                        addToLog(logFile,"> NOT IMPORTED (" + FileName + ") CLASS NOT FOUND : " + ClassID);
                        continue;
                        //throw new Exception("ClassID not found for name : " + FolderName);
                    }
                    log.info("GIBDoc Import Found Doc Class ID (" + ClassID + ") for : " + FolderName);
                    IDocument GIBDoc = null;
                    GIBDoc = getGIBDoc(ClassID,FileName);
                    if(GIBDoc != null){
                        log.info("GIB DOC EXIST : " + FileName);
                        addToLog(logFile,"> ALREADY IMPORTED (" + FileName + ") : " + GIBDoc.getID());
                        continue;
                    }
                    IDocument doc = newFileToDocumentClass(filePath, ClassID);
                    log.info("GIBDoc Created");
                    doc.setDescriptorValue("_CIF",CIFNumber);
                    doc.setDescriptorValue("_AccountNumber",AccountNumber);
                    doc.setDescriptorValue("_DocumentDate",DocumentDate);
                    doc.setDescriptorValue("_StoredDate",StoredDate);
                    doc.setDescriptorValue("_Title",FileName);
                    doc.setDescriptorValue("_DocumentType", DocumentType);
                    doc.commit();
                    importedCount++;
                    addToLog(logFile,"> IMPORTED (" + FileName + ") to " + CustomerType + " / " + FolderName);
                    log.info("GIBDoc Imported Doc ID : " + doc.getID());
                }
                this.log.info("Import GIBDoc from Excel Finished");
                addToLog(logFile,"IMPORTED DOCS [" + importedCount + " / " + totalCount + "]");
                return this.resultSuccess("Ended successfully");
            } catch (Exception e) {
                this.log.error("Exception Caught");
                this.log.error(e.getMessage());
                return resultError(e.getMessage());
            }
        }
    }
    private void addToLog(File file, String log) throws Exception {
        try {
            FileWriter fileWriter = new FileWriter(file, true);
            BufferedWriter bufferedWriter = new BufferedWriter(fileWriter);
            bufferedWriter.write(log);
            bufferedWriter.newLine();
            bufferedWriter.close();
        }catch (Exception e){
            throw new Exception("Write error to log file : " + e);
        }
    }
    private void writeLogToFile(List<String> importLog, String fileName) {
        Path logFilePath = Paths.get(fileName);
        try {
            Files.write(logFilePath, importLog);
            log.info("Log writed: " + fileName);
        } catch (IOException e) {
            log.error("Log not writed: " + e.getMessage());
        }
    }
    public String getCategoryIDFromGVlist(String name) throws Exception {
        String rtrn = "";
        IStringMatrix settingsMatrix = srv.getStringMatrixByID("DocumentClasses", ses);
        if(settingsMatrix!=null) {
            for (int i = 0; i < settingsMatrix.getRowCount(); i++) {
                String rowID = settingsMatrix.getValue(i, 0);
                String rowName = settingsMatrix.getValue(i, 1);
                if (rowName.equalsIgnoreCase(name)) {
                    rtrn = rowID;
                    break;
                }
            }
        }
        return rtrn;
    }
    public String getDocumentTypeFromGVList(String key1, String key2) {
        String rtrn = "";
        String rowCustomerType = "";
        String rowFolderName = "";
        String rowDocumentType = "";
        String rowOldDocumentType = "";
        for(int i = 0; i < settingsMatrix.getRowCount(); i++) {
            rowCustomerType = settingsMatrix.getValue(i, 0);
            rowFolderName = settingsMatrix.getValue(i, 1);
            rowDocumentType = settingsMatrix.getValue(i, 2);
            rowOldDocumentType = settingsMatrix.getValue(i, 4);

            //if (!Objects.equals(rowValuePrjCode, CIFNumber)){continue;}
            if (!Objects.equals(rowOldDocumentType, key1)){continue;}
            if (!Objects.equals(rowFolderName, key2)){continue;}
            //if (!Objects.equals(rowValueParamMyComp, "1")){continue;}

            rtrn = rowDocumentType;
            break;
        }
        return rtrn;
    }
    public String getFolderNameFromGVList(String key1, String key2) {
        String rtrn = "";
        String rowCustomerType = "";
        String rowFolderName = "";
        String rowDocumentType = "";
        String rowOldDocumentType = "";
        for(int i = 0; i < settingsMatrix.getRowCount(); i++) {
            rowCustomerType = settingsMatrix.getValue(i, 0);
            rowFolderName = settingsMatrix.getValue(i, 1);
            rowDocumentType = settingsMatrix.getValue(i, 2);
            rowOldDocumentType = settingsMatrix.getValue(i, 4);

            //if (!Objects.equals(rowValuePrjCode, CIFNumber)){continue;}
            if (!Objects.equals(rowOldDocumentType, key1)){continue;}
            if (!Objects.equals(rowCustomerType, key2)){continue;}
            //if (!Objects.equals(rowValueParamMyComp, "1")){continue;}

            rtrn = rowFolderName;
            break;
        }
        return rtrn;
    }
    public static HashMap<String, Row> listOfDocuments(Workbook workbook) throws IOException {
        HashMap<String, Row> rtrn = new HashMap();
        Sheet sheet = workbook.getSheetAt(0);
        Iterator var3 = sheet.iterator();

        while(var3.hasNext()) {
            Row row = (Row)var3.next();
            if (row.getRowNum() != 0 && row.getRowNum() != 1) {
                Cell cll1 = row.getCell(0);
                if (cll1 != null) {
                    String indx = cll1.getRichStringCellValue().getString();
                    if (!indx.equals("") && !indx.equals("File Name")) {
                        rtrn.put(indx, row);
                    }
                }
            }
        }
        return rtrn;
    }
    public static List<String> fieldsOfDocuments(Workbook workbook) throws IOException {
        List<String> rtrn = new ArrayList<>();
        Sheet sheet = workbook.getSheetAt(0);
        String indx = "";
        Iterator var3 = sheet.iterator();
        Integer c = 0;
        while(var3.hasNext()) {
            Row row = (Row)var3.next();
            if (row.getRowNum() == 0) {
                while (c<35) { ///toplam fields sayısı
                    Cell cll1 = row.getCell(c);
                    if (cll1 != null) {
                        indx = cll1.getRichStringCellValue().getString();
                        if(Objects.equals(indx, "")){break;}
                        rtrn.add(indx);
                    }else{break;}
                    c++;
                }
            } else{break;}
        }
        return rtrn;
    }

    public String exportDocumentContent(IDocument document, String exportPath) throws IOException {
        String expt = "";
        String documentID = document.getDocumentID().getID();
        documentID = documentID.replaceAll(":", ".");

        for(int representationConter = 0; representationConter < document.getRepresentationCount(); ++representationConter) {
            for(int partDocumentCounter = 0; partDocumentCounter < document.getPartDocumentCount(representationConter); ++partDocumentCounter) {
                IDocumentPart partDocument = document.getPartDocument(representationConter, partDocumentCounter);
                InputStream inputStream = partDocument.getRawDataAsStream();

                try {
                    IFDE fde = partDocument.getFDE();
                    if (fde.getFDEType() == 3) {
                        expt = exportPath + "/output_" + documentID + "." + ((IFileFDE)fde).getShortFormatDescription();
                        FileOutputStream fileOutputStream = new FileOutputStream(expt);

                        try {
                            byte[] bytes = new byte[2048];

                            int length;
                            while((length = inputStream.read(bytes)) > -1) {
                                fileOutputStream.write(bytes, 0, length);
                            }
                        } catch (Throwable var15) {
                            try {
                                fileOutputStream.close();
                            } catch (Throwable var14) {
                                var15.addSuppressed(var14);
                            }

                            throw var15;
                        }

                        fileOutputStream.close();
                    }
                } catch (Throwable var16) {
                    if (inputStream != null) {
                        try {
                            inputStream.close();
                        } catch (Throwable var13) {
                            var16.addSuppressed(var13);
                        }
                    }

                    throw var16;
                }

                if (inputStream != null) {
                    inputStream.close();
                }
            }
        }

        return expt;
    }
    public static Object getValue(Cell cell, CellType type) {
        switch (type) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return getLocalDateTime(cell.getDateCellValue().toString());
                } else {
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value)) {
                        return (long)value;
                    }

                    return value;
                }
            case STRING:
                return cell.getStringCellValue();
            case FORMULA:
                return getValue(cell, cell.getCachedFormulaResultType());
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case _NONE:
                return null;
            case BLANK:
                return null;
            case ERROR:
                return null;
            default:
                return null;
        }
    }
    public static LocalDateTime getLocalDateTime(String strDate) {
        strDate = strDate.replace("TRT", "Europe/Istanbul");
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEE MMM dd HH:mm:ss zzzz yyyy");
        ZonedDateTime zdt = ZonedDateTime.parse(strDate, formatter);
        LocalDateTime ldt = zdt.toLocalDateTime();
        return ldt;
    }
    public IArchiveFolderClass createNewMDPRecord(){
        try{
            log.info("Creating new MDP record..");
            IMetaDataManager dataManager = srv.getMetaDataManager(getSes());
            IArchiveFolderClassModifiable archiveFolderClassModifiable = dataManager.createArchiveFolderClass(Constants.ClassIDs.GIBRecord);
            archiveFolderClassModifiable.commit();
            //ISerClassFactory classFactory = srv.getClassFactory();
            //IArchiveClass aClass = srv.getArchiveClass(id,getSes());
            //rtrn = classFactory.getDocumentInstance(getSes(), aClass, getSes().getDatabaseByName("PRJ_FOLDER"), null);
            return (IArchiveFolderClass) archiveFolderClassModifiable;
        }catch(Exception e){
            log.error("Exception caught...createNewMDPRecord..:" + e.getMessage());
            return null;
        }
    }
    private IFolder newGIBRecord() throws Exception{
        try{
            IFolderConnection folderConnection = ses.getFolderConnection();
            IFolder rtrn = folderConnection.createFolder();
            IArchiveFolderClass afc = srv.getArchiveFolderClass(Constants.ClassIDs.GIBRecord, ses);
            IDatabase db = ses.getDatabase(afc.getDefaultDatabaseID());
            rtrn.init(afc);
            rtrn.setDatabaseName(db.getDatabaseName());
            return rtrn;
        }catch(Exception e){
            log.error("Exception caught...newGIBDoc..:" + e.getMessage());
            return null;
        }
    }
    public IDocument getGIBDoc(String classID, String Filename)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(classID).append("'")
                .append(" AND ")
                .append(Constants.Literals.TITLE).append(" = '").append(Filename).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = createQuery(new String[]{Constants.Databases.GIBDocDB} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IDocument) informationObjects[0];
    }
    public IDocument getGIBDocOLD(String CIFNo, String AccNo, String CustomerType, String DocType)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Constants.ClassIDs.GIBRecord).append("'")
                .append(" AND ")
                .append(Constants.Literals.CIF_NUMBER).append(" = '").append(CIFNo).append("'")
                .append(" AND ")
                .append(Constants.Literals.ACC_NUMBER).append(" = '").append(AccNo).append("'")
                .append(" AND ")
                .append(Constants.Literals.CUSTOMER_TYPE).append(" = '").append(CustomerType).append("'")
                .append(" AND ")
                .append(Constants.Literals.DOC_TYPE).append(" = '").append(DocType).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = createQuery(new String[]{Constants.Databases.GIBRecordDB} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IDocument) informationObjects[0];
    }
    public IInformationObject[] createQuery(String[] dbNames , String whereClause , int maxHits){
        String[] databaseNames = dbNames;

        ISerClassFactory fac = getSrv().getClassFactory();
        IQueryParameter que = fac.getQueryParameterInstance(
                getSes() ,
                databaseNames ,
                fac.getExpressionInstance(whereClause) ,
                null,null);
        if(maxHits > 0) {
            que.setMaxHits(maxHits);
            que.setHitLimit(maxHits + 1);
            que.setHitLimitThreshold(maxHits + 1);
        }
        IDocumentHitList hits = que.getSession() != null? que.getSession().getDocumentServer().query(que, que.getSession()):null;
        if(hits == null) return null;
        else return hits.getInformationObjects();
    }
    public static String getValueFromRow(List<String> fields,Row row, String fieldName) throws IOException {
        String rtrn = "";

        int c = 0;
        for (String field :fields) { ///toplam fields sayısı
            if(Objects.equals(fieldName, field)) {
                Cell cll1 = row.getCell(c);
                if (cll1 != null) {
                    rtrn = cll1.getRichStringCellValue().getString();
                    break;
                }
            }
            c++;
        }
        return rtrn;
    }
    public IDocument newGIBDoc(String tpltSavePath) throws Exception {
        IDocument doc = newFileToDocumentClass(tpltSavePath, Constants.ClassIDs.GIBDoc);
        doc.setDescriptorValue("ccmFileName" , "Deleted Process Log File.xlsx");
        doc.setDescriptorValue("ccmPrjDocDocType" , "DOC");
        doc.setDescriptorValue("ccmReferenceNumber" , getEventTask().getProcessInstance().getMainInformationObjectID());
        doc.commit();
        return doc;
    }
    public IDocument newFileToDocumentClass(String filePath, String archiveClassID) throws Exception {
        IArchiveClass cls = srv.getArchiveClass(archiveClassID, ses);
        if (cls == null) cls = srv.getArchiveClassByName(ses, archiveClassID);
        if (cls == null) throw new Exception("Document Class: " + archiveClassID + " not found");

        String dbName = ses.getDatabase(cls.getDefaultDatabaseID()).getDatabaseName();
        IDocument doc = srv.getClassFactory().getDocumentInstance(dbName, cls.getID(), "0000", ses);

        File file = new File(filePath);
        IRepresentation representation = doc.addRepresentation(".pdf" , "Signed document");
        IDocumentPart newDocumentPart = representation.addPartDocument(filePath);
        //doc.commit();
        return doc;
    }

}
