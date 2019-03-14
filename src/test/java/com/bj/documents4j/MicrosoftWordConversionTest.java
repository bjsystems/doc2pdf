package com.bj.documents4j;

import static org.junit.Assert.assertEquals;

import java.util.Arrays;
import java.util.Collection;

import org.junit.BeforeClass;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;

import com.documents4j.api.DocumentType;
import com.documents4j.conversion.msoffice.AbstractMicrosoftOfficeConversionTest;
import com.documents4j.conversion.msoffice.Document;
import com.documents4j.conversion.msoffice.DocumentTypeProvider;
import com.documents4j.conversion.msoffice.MicrosoftWordBridge;


import static com.bj.documents4j.MicrosoftWordDocument.*;

@RunWith(Parameterized.class)
public class MicrosoftWordConversionTest extends AbstractMicrosoftOfficeConversionTest {

    public MicrosoftWordConversionTest(Document valid,
                                       Document corrupt,
                                       Document inexistent,
                                       DocumentType sourceDocumentType,
                                       DocumentType targetDocumentType,
                                       String targetFileNameSuffix,
                                       boolean supportsLockedConversion) {
        super(new DocumentTypeProvider(valid, corrupt, inexistent, sourceDocumentType, targetDocumentType, targetFileNameSuffix, supportsLockedConversion));
    }

    @Parameterized.Parameters
    public static Collection<Object[]> data() {
        return Arrays.asList(new Object[][]{
                {DOCX_VALID, DOCX_CORRUPT, DOCX_INEXISTENT, DocumentType.MS_WORD, DocumentType.PDF, "pdf", true}
        });
    }

    @BeforeClass
    public static void setUpConverter() throws Exception {
        AbstractMicrosoftOfficeConversionTest.setUp(MicrosoftWordBridge.class, MicrosoftWordScript.ASSERTION, MicrosoftWordScript.SHUTDOWN);
    }
}
