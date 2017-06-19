require "zip/zip"
require "xml"
 
module OOXML
  
    # Defines a basic ooxml document, that creates a handle to randomly read and access the low-level contents of a OOXML file.
    class Document
        
        @core_props = DocumentProperties.new
        @content_types = Array(ContentType).new
        @relations = Array(DocumentRelation).new
        
        def initialize(@file : File)
            @zip = Zip::File.new(@file)
            read_content_types
            read_relations
            read_core_props
        end
        
        def initialize(file : String)
            initialize(File.new(file))
        end
        
        def io() @file; end
            
        def zip() @zip; end
            
        def close()
            @file.close
        end
        
        private def read_core_props()
            entry_name = file_for_relation?(OOXML::RELATION_CORE_PROPERTIES)
            if(entry_name)
                open_xml entry_name do |xml|
                    @core_props = DocumentProperties.new(xml)
                end
            else
                raise OOXML::OOXMLFormatException.new("Could not load core properties")
            end
        end
        
        private def read_content_types()
            open_xml "[Content_Types].xml" do |xml|
                xml.children.select(&.element?).each do |child|
                    if child.name == "Default"
                        @content_types << ContentType.new(child["Extension"], nil, child["ContentType"])
                    elsif child.name == "Override"
                        @content_types << ContentType.new(nil, child["PartName"], child["ContentType"])
                    end
                end
            end
        end
        
        private def read_relations()
            open_xml "_rels/.rels" do |xml|
                xml.children.select(&.element?).each do |child|
                    if child.name == "Relationship"
                        @relations << DocumentRelation.new(child["Id"], child["Type"], child["Target"])
                    end
                end
            end
        end
        
        def core_props() @core_props; end
        def content_types() @content_types; end
            
        def open_xml(name : String)
            entry = @zip[name]
            if entry
                entry.open do |data|
                    XML.parse(data).first_element_child
                end
            else
                raise OOXML::OOXMLFormatException.new("Could not find file #{name}")
            end
        end
        
        def open_xml(name : String)
            entry = @zip[name]
            if entry
                entry.open do |data|
                    elem = XML.parse(data).first_element_child
                    if elem
                        yield elem
                    else
                        raise OOXML::OOXMLFormatException.new("File #{name} is empty")
                    end
                end
            else
                raise OOXML::OOXMLFormatException.new("Could not find file #{name}")
            end
        end
        
        def content_type_of?(filename : String)
            if !filename.starts_with?('/')
                return content_type_of?("/#{filename}")
            end
            highest_match = 0
            found = nil
            @content_types.each do |content_type|
                if (match = content_type.matches(filename)) > highest_match
                    found = content_type.content_type
                    highest_match = match
                end
            end
            found
        end
        
        def file_for_relation?(relation_type : String)
            @relations.each do |relation|
                if relation.type == relation_type
                    return relation.target
                end
            end
            nil
        end
        
    end
    
    class DocumentProperties
        
        @creator : String | Nil = nil
        @lastModifiedBy : String | Nil = nil
        @modified : Time | Nil = nil
        @created : Time | Nil = nil
       
        def initialize(xml : XML::Node)
            assertEquals "coreProperties", xml.name
            namespace = xml.namespace
            if !namespace
                raise OOXML::OOXMLAssertionException.new("cp", nil) 
            end
            assertEquals "cp", namespace.prefix
            
            xml.children.select(&.element?).each do |child|
                if child.name == "creator"
                    @creator = child.content
                elsif child.name == "lastModifiedBy"
                    @lastModifiedBy = child.content
                elsif child.name == "created"
                    @created = Time::Format::ISO_8601_DATE_TIME.parse(child.content)
                elsif child.name == "modified"
                    @modified = Time::Format::ISO_8601_DATE_TIME.parse(child.content)    
                end
            end
            
            assertNotNil(@creator)
            assertNotNil(@lastModifiedBy)
            assertNotNil(@modified)
            assertNotNil(@created)
        end
        
        def initialize(); end
        
        def creator() @creator; end
        def creator=(creator : String) @creator = creator; end
        
        def lastModifiedBy() @lastModifiedBy; end
        def lastModifiedBy=(lastModifiedBy : String) @lastModifiedBy = lastModifiedBy; end
        
        def modified() @modified; end
        def modified=(modified : Time) @modified = modified; end
        
        def created() @created; end
        def created=(created : Time) @created = created; end
    end
    
    class DocumentRelation
       
        def initialize(@id : String, @type : String, @target : String); end
            
        def type() @type; end
        def target() @target; end
        
    end
    
    class ContentType
        
        def initialize(@extension : String | Nil, @part_name : String | Nil, @content_type : String); end
            
        def matches(filename : String)
            if part_name = @part_name
                if part_name == filename
                    return 2
                end
            end
            if extension = @extension
                if filename.ends_with?(extension)
                    return 1
                end
            end
            0
        end
        
        def content_type() @content_type; end
    end
  
end

def assertEquals(expected : String, actual : String | Nil)
    if !actual
        raise OOXML::OOXMLAssertionException.new(expected, actual)
    end
    if expected != actual
        raise OOXML::OOXMLAssertionException.new(expected, actual)
    end
end

def assertNotNil(actual : String | Nil) 
    if !actual
        raise OOXML::OOXMLAssertionException.new("any string", nil)
    end
end

def assertNotNil(actual : Time | Nil)
    if !actual
        raise OOXML::OOXMLAssertionException.new("any time", nil)
    end
end
