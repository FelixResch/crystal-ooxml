require "./document"

module OOXML
   
    class Workbook
        
        @worksheets = Array(Worksheet).new
        @relations = Array(WorkbookRelation).new
        @base_path = ""
        
        def initialize(@document : OOXML::Document)
            file : String | Nil = @document.file_for_relation?(OOXML::RELATION_OFFICE_DOCUMENT)
            if !file
                raise OOXML::OOXMLWorkbookException.new("No main office document found")
            end
            @file = file.as(String)
            content_type = @document.content_type_of?(@file)
            if content_type != OOXML::CONTENT_TYPE_WORKBOOK
                raise OOXML::OOXMLWorkbookException.new("Document is not a workbook. Found: #{content_type}")
            end
            @base_path = @file.partition('/')[0]
            read_relations
            read_workbook
        end
        
        private def read_relations()
            parts = @file.rpartition('/')
            rels_name = parts[0] + "/_rels/" + parts[2] + ".rels"
            @document.open_xml rels_name do |xml|
                xml.children.select(&.element?).each do |child|
                    if child.name == "Relationship"
                        @relations << WorkbookRelation.new(child["Id"], child["Type"], child["Target"])
                    end
                end
            end
        end
        
        private def read_workbook()
            @document.open_xml @file do |xml|
                xml.children.select(&.element?).each do |child|
                    if child.name == "sheets"
                        read_sheets child
                    end
                end
            end
        end
        
        private def read_sheets(xml : XML::Node)
            xml.children.select(&.element?).each do |child|
                @worksheets << Worksheet.new(child["name"], child["sheetId"], child["state"]?, child["id"], self)
            end
        end
        
        def close()
            @document.close
        end
        
        def file_by_relation_id(relation_id : String)
            @relations.each do |relation|
                if(relation.id == relation_id)
                    return relation
                end
            end
            nil
        end
        
        def open_xml(filename : String)
            @document.open_xml(@base_path + "/" + filename)
        end
        
        def open_xml(filename : String, &block)
            @document.open_xml(@base_path + "/" + filename, &block)
        end
        
        def content_type_of?(filename : String)
            @Document.content_type_of?(@base_path + "/" + filename)
        end
        
    end
    
    class WorkbookRelation
       
        def initialize(@id : String, @type : String, @target : String); end
            
        def id() @id; end
    end
    
    class Worksheet
        
        def initialize(@name : String, @sheetId : String, @state : String | Nil, @relId : String, @workbook : Workbook)
            read_data
        end
        
        private def read_data()
            @workbook.file_by_relation_id(@relId) do |relation|
                if relation.type != OOXML::RELATION_WORKSHEET
                    raise OOXML::OOXMLWorkbookException.new("Invalid relationship type for worksheet")
                end
                if !((content_type = @workbook.content_type_of?(relation.target)) && content_type == OOXML::CONTENT_TYPE_WORKSHEET)
                    raise OOXML::OOXMLWorkbookException.new("Invalid content_type for worksheet")
                end
                @workbook.open_xml relation.target do |xml|
                    puts xml.to_s
                end
            end
        end
    end
end
