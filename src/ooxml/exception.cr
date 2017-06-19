 
module OOXML

    class OOXMLFormatException < Exception
        
    end
    
    class OOXMLAssertionException < Exception
        
        def initialize(expected : String, actual : String | Nil)
            initialize("Assertion failed! Expected: <#{expected}>, Got: <#{actual}>")
        end
    end
    
    class OOXMLWorkbookException < Exception
        
    end
end
