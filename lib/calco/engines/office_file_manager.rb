require 'zip'
require 'tmpdir'
require 'pathname'
require 'stringio'
require 'tempfile'
require 'rexml/document'

module Calco

  class OfficeFileManager
    
    Descriptor = Struct.new(:stream, :file, :new_sheet, :header)
    
    def initialize ods_template
      
      @ods_template = ods_template
      
      @sheets = Hash.new do |h, k|
      
        temp_file = Tempfile.new("office-gen-sheet-")

        stream = open(temp_file, 'w:utf-8')
        
        h[k] = Descriptor.new(stream, temp_file, false)
        
      end
      
    end
    
    def save definitions, to_filename
      
      @sheets.each do |name, descriptor|
      
        descriptor.stream.close
      
      end
      
      flush_content definitions, to_filename
      
    end
    
    def add_empty_row sheet
      @sheets[sheet.sheet_name].stream.write '<table:table-row/>'
    end
    
    def add_row sheet
    
      stream = @sheets[sheet.sheet_name].stream
      
      stream.write '<table:table-row>'

      yield stream

      stream.write '</table:table-row>'
      
    end
    
    private
    
    def flush_content definitions, to_filename

      content_xml_file = Tempfile.new('office-gen')
      result_xml_file = Tempfile.new('office-gen')
      
      extract_template_content content_xml_file

      prepare_content_file definitions, content_xml_file, result_xml_file

      create_file result_xml_file, to_filename
      
    end
    
    def extract_template_content content_xml_file
      
      Zip::File.open(@ods_template) do |zipfile|
        content = zipfile.read("content.xml")
        open(content_xml_file, "w") {|out| out.write content}
      end
      
    end
    
    def prepare_content_file definitions, content_xml_file, result_xml_file

      file = File.new(content_xml_file)

      xml = REXML::Document.new(file)

      prepare_xml(xml)

      temp_file = create_temporary(xml, result_xml_file)

      File.open(result_xml_file, 'w') do |stream|

        @out_stream = stream

        File.open(temp_file, 'r').each do |line|

          if line =~ /\A%%%Insert data here \[(.*)\]%%%\Z/
            
            write_sheet $1, definitions
            
          elsif line =~ /(.*)(<\/office:spreadsheet>)(.*)/
          
            @out_stream.write $1
            
            write_new_sheets
          
            @out_stream.write $2
            @out_stream.write $3
            
          else
            
            @out_stream.write line
            
          end

        end

      end

    end
    
    def create_file result_xml_file, to_filename
    
      FileUtils.cp(@ods_template, to_filename)

      Zip::File.open(to_filename) do |zipfile|

        zipfile.get_output_stream("content.xml") do |os|

          File.open(result_xml_file).each_line do |line|
            os.puts line
          end

        end

      end

    end
    
    def prepare_xml xml_doc

      root = xml_doc.root

      @sheets.each_key do |name|
      
        table = root.elements["//table:table[@table:name='#{name}']"]
        
        unless table
          @sheets[name].new_sheet = true
          next
        end
        
        state = :waiting_header

        table.each_element('table:table-row') do |row|

          @sheets[name].header = stringify(row) if state == :waiting_header

          state = :header_consumed

          table.delete_element(row)
            
        end
        
        table.add_text marker(name)

      end

    end

    def write_sheet name, definitions
      
      sheet = @sheets[name]
      
      if definitions[name].has_titles?
      
        if sheet.header
          @out_stream.write sheet.header
        else
          $sdterr.puts "Cannot find template row in #{@ods_template} for #{name}"
        end
        
      end
      
      @out_stream.write sheet.file.read
      
    end
    
    def write_new_sheets
      
      @sheets.each do |name, descriptor|
        
        next unless descriptor.new_sheet
        
        @out_stream.write "<table:table table:name='#{name}'>"
        @out_stream.write descriptor.file.read
        @out_stream.write '</table:table>'
        
      end
      
    end
    
    def create_temporary xml, to_filename

      to = Pathname.new(to_filename)

      temp_file = Tempfile.new('office-gen', to.dirname.to_s)

      File.open(temp_file, 'w') { |stream| stream.puts xml }

      temp_file

    end

    def marker(name)
      "\n%%%Insert data here [#{name}]%%%\n"
    end
    
    def stringify(row)
    
      buffer = StringIO.new
    
      REXML::Formatters::Default.new.write(row, buffer)
      
      buffer.string
      
    end
    
  end
  
end
