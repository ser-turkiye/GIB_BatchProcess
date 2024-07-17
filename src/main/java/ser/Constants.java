package ser;

public class Constants {

    public static class Templates{
        public static class CRS{
            public static final String OriginalTemplate = "C:\\SER\\TEMP\\Templates\\CRSTemplate.docx";
            public static final String OriginalTemplateXlsx = "C:\\SER\\TEMP\\Templates\\CRSTemplate.xlsx";
            public static final String HolderPath = "C:\\SER\\TEMP\\output-1.docx";
            public static String FinalPath = "C:\\SER\\TEMP\\output-2.docx";
            public static String FinalPathXlsx = "C:\\SER\\TEMP\\output-2.xlsx";
        }
        public static class Transmittal{
            public static final String OriginalTemplate = "C:\\SER\\TEMP\\Templates\\TransmittalTemplate.docx";
            public static final String HolderPath = "C:\\SER\\TEMP\\output-3.docx";
            public static String FinalPath = "C:\\SER\\TEMP\\output-4.docx";
        }
    }
    public static class Databases{
        public static final String GIBRecordDB = "GIB_CUSTOMER";
        public static final String GIBDocDB = "GIB_DOCS";
    }
    public static class ClassIDs{
        public static final String GIBRecord = "23368685-1da0-4fa8-90a5-441048ed4f2e";
        public static final String GIBDoc = "23368685-1da0-4fa8-90a5-441048ed4f2e";
        public static final String ImportDoc = "dd637787-b26b-4714-a68a-96f9ea754b40";
    }
    public static class Descriptors {
        public static final String ProjectNumber = "ProjectNumber";
        public static final String NumberReference= "ObjectNumberReference";
        public static final String Discipline = "OrgDepartment";
    }

    public static class Literals {
        public static final String CIF_NUMBER = "D_CIF";
        public static final String ACC_NUMBER = "ACCOUNTNUMBER";
        public static final String CUSTOMER_TYPE = "D_CUSTOMERTYPE";
        public static final String DOC_TYPE = "D_DOCUMENTTYPE";
        public static final String TITLE = "D_TITLE";
        public static final String ProjectFolderDB = "PRJ_FOLDER";
    }

}
