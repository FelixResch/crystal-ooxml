require "./spec_helper"

WORKBOOK_MS = "blank.xlsx"
WORKBOOK_LIBRE = "blank_libre.xlsx"

DOC_LIBRE = "blank_libre.docx"

describe OOXML do
  
    describe OOXML::Document do
        
        it "correctly opens a file when a string is provided" do
            document = OOXML::Document.new(WORKBOOK_MS)
            document.io.is_a?(File).should be_true
            document.close
        end
        
        it "uses the IO provided by constructor call" do
            document = OOXML::Document.new(File.new(WORKBOOK_MS))
            document.io.is_a?(File).should be_true
            document.close
        end
        
        it "test core properties correct" do
            document = OOXML::Document.new(File.new(WORKBOOK_MS))
            document.core_props.should be_truthy
            document.core_props.creator.should eq "Felix Resch"
            document.core_props.lastModifiedBy.should eq "Felix Resch"
            document.close
        end
        
        it "find correct content-types" do
            document = OOXML::Document.new(WORKBOOK_MS)
            document.content_type_of?("text.xml").should eq "application/xml"
            document.content_type_of?("/xl/workbook.xml").should eq "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            document.close
        end
        
        it "find correct file for relation" do
            document = OOXML::Document.new(WORKBOOK_MS)
            document.file_for_relation?("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument").should eq "xl/workbook.xml"
            document.close
        end
        
        it "open MS Word document" do
            document = OOXML::Document.new(DOC_LIBRE)
            document.close
        end
        
        it "open MS Word document as workbook" do
            expect_raises OOXML::OOXMLWorkbookException do
                workbook = OOXML::Workbook.new(OOXML::Document.new(DOC_LIBRE))
            end
        end
    end
    
    describe OOXML::Workbook do
        
        it "opens a correct workbook without throwing" do
            workbook = OOXML::Workbook.new(OOXML::Document.new(WORKBOOK_MS))
            workbook.close
            workbook = OOXML::Workbook.new(OOXML::Document.new(WORKBOOK_LIBRE))
            workbook.close
        end
        
    end
    
end
