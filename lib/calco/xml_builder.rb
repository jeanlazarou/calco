module Calco

  class XMLBuilder

    def initialize tag
      @tag = tag
      @kids = {}
      @attributes = {}
    end
    
    # +attributes+ is a hash with key/value pairs, keys are the attribute names
    def << attributes
    
      attributes.each do |attribute, value|
        @attributes[attribute] = escape(value)
      end
      
    end

    def attribute name, value
      @attributes[name] = escape(value)
    end
    
    # add a child text-node with the given +tag+/+value+
    def add_child tag, value
      @kids[tag] = value
    end
    
    def build
    
      buffer = "<#{@tag} "
      
      return buffer + '/>' if @kids.empty? && @attributes.empty?
      
      @attributes.each do |attribute, value|
        buffer << "#{attribute}='#{value}' "
      end
      
      if @kids.empty?
        buffer << '/>'
      else
        
        buffer << '>'
        
        @kids.each do |text_tag, value|
          buffer << "<#{text_tag}><![CDATA[#{value}]]></#{text_tag}>"
        end
      
        buffer << "</#{@tag}>"
      
      end
      
      buffer
      
    end
    
    def escape str
      
      str = str.to_s
      
      str.gsub("'", '&apos;')
      
    end
    
  end
  
end
