require 'calco/xml_builder'

module Calco

  describe XMLBuilder do

    it "generates empty tag" do
    
      builder = XMLBuilder.new('hello')
      
      expect(builder.build).to eq('<hello />')
    
    end
    
    it "generates tag + one attribute" do
    
      builder = XMLBuilder.new('hello')
      
      builder << {:name => 'Max'}
      
      expect(builder.build).to eq("<hello name='Max' />")
    
    end
    
    it "generates tag + many attributes" do
    
      builder = XMLBuilder.new('hello')
      
      builder << {name: 'Max', greeting: 'Bonjour', today: Date.new(2014, 3, 16)}
      
      expect(builder.build).to eq("<hello name='Max' greeting='Bonjour' today='2014-03-16' />")
    
    end
    
    it "generates tag + many attributes (alternate notation)" do
    
      builder = XMLBuilder.new('hello')
      
      builder.attribute :name, 'Max'
      builder.attribute :greeting, 'Bonjour'
      builder.attribute "today", Date.new(2014, 3, 16)
      
      expect(builder.build).to eq("<hello name='Max' greeting='Bonjour' today='2014-03-16' />")
    
    end
    
    it "overwrites attributes" do
    
      builder = XMLBuilder.new('hello')
      
      builder.attribute :name, 'Max'

      builder << {:name => 'Joe'}
      
      expect(builder.build).to eq("<hello name='Joe' />")
    
    end
    
    it "generates tag + attributes + text child nodes" do
    
      builder = XMLBuilder.new('hello')
      
      builder << {name: 'Max', town: 'London'}
      
      builder.add_child "comment", "He's Joe's best friend"
      builder.add_child "info", "He lives somewhere"
      
      expect(builder.build).to eq(
                 "<hello name='Max' town='London' >" + 
                     "<comment><![CDATA[He's Joe's best friend]]></comment>" + 
                     '<info><![CDATA[He lives somewhere]]></info>' + 
                  '</hello>'
              )
    
    end
    
    it "escapes attibute values" do
      
      builder = XMLBuilder.new('hello')
      
      builder << {name: 'Max', last_name: "O'Neil"}
      
      expect(builder.build).to eq("<hello name='Max' last_name='O&apos;Neil' />")
      
    end
    
  end

end
