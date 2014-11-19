require 'date'

require 'zip'
require 'tmpdir'
require 'pathname'
require 'tempfile'
require 'rexml/document'

require 'calco'
require 'calco/xml_builder'

module Calco

  class OfficeEngine < DefaultEngine

    def initialize ods_template, first_row_is_header = true
      @ods_template = ods_template
      @first_row_is_header = first_row_is_header
    end
    
    # output is a String (as a file name)
    def save doc, to_filename, &data_iterator

      content_xml_file = Tempfile.new('office-gen')
      result_xml_file = Tempfile.new('office-gen')
      
      Zip::File.open(@ods_template) do |zipfile|
        content = zipfile.read("content.xml")
        open(content_xml_file, "w") {|out| out.write content}
      end

      write_result_content doc, content_xml_file, result_xml_file, @first_row_is_header, &data_iterator

      FileUtils.cp(@ods_template, to_filename)

      Zip::File.open(to_filename) do |zipfile|

        zipfile.get_output_stream("content.xml") do |os|

          File.open(result_xml_file).each_line do |line|
            os.puts line
          end

        end

      end

    end

    def empty_row
      @out_stream.write '<table:table-row/>'
    end
    
    def write_row sheet, row_id

      return if row_id == 0 && sheet.has_titles?
      
      row_id += 1 # office sheet indexes start at 1
      
      cells = sheet.row(row_id)

      @out_stream.write '<table:table-row>'

      cells.each_index do |i|

        cell = cells[i]

        @out_stream.write cell

      end

      @out_stream.write '</table:table-row>'

    end

    def generate_cell row_number, column, cell_style, column_style, column_type

      return '<table:table-cell/>' unless column
      return '<table:table-cell/>' if column.absolute_row && column.absolute_row != row_number

      currency = nil
      
      cell = column.generate(row_number)

      cell_style = cell_style.generate(row_number) if cell_style

      if column_type

        if column_type == '%'
          column_type = 'percentage'
        elsif column_type =~ /\$([A-Z]{3})/
          currency = "#{$1}"
          column_type = 'currency' 
        end

      end

      office_cell column, column_style, column_type, currency, cell.to_s, cell_style

    end

    def column_reference element, row
      element.is_a?(Current) ? "#{super}" : "[.#{super}]"
    end

    def value element
    
      if ! element.respond_to?(:value)
        s = element.to_s
      else
        s = element.value.to_s
      end
      
      s.gsub!('"', '&quot;') if s =~ /"/
      s
      
    end

    def style statement, row
      "&amp;T(ORG.OPENOFFICE.STYLE(#{statement.generate(row)}))"
    end

    def operator op

      if op == '!='
        '&lt;&gt;'
      elsif op == '<'
        '&lt;'
      elsif op == '>'
        '&gt;'
      else
        op
      end

    end

    def current
      'ORG.OPENOFFICE.CURRENT()'
    end

    private

    def office_cell column, column_style, column_type, currency, value, cell_style

      xml = XMLBuilder.new('table:table-cell')
      
      if column.is_a?(Formula)

        column_type = 'float' unless column_type

        xml << {'table:formula' => "of:=#{value}#{cell_style}"}

      elsif column.respond_to?(:value)
      
        if column.value.nil?
        
          return xml.build
        
        elsif column.value.is_a?(Numeric)

          column_type = 'float' unless column_type

          xml << {'office:value' => "#{value}#{cell_style}"} unless cell_style

        elsif column.value.is_a?(Date)

          column_type = 'date' unless column_type

          xml << {'office:date-value' => "#{value}#{cell_style}"} unless cell_style

        else

          column_type = 'string' unless column_type

          if cell_style
            value = office_string_escape(value)
          else
            xml.add_child 'text:p', office_string_value(value)
          end
          
        end

        if cell_style
          xml << {'table:formula' => "of:=#{value}#{cell_style}"}
        end
        
      end

      xml << {'table:style-name'  => column_style}   if column_style
      xml << {'office:currency'   => currency}       if currency
      xml << {'office:value-type' => column_type}
      
      xml.build
      
    end

    def office_string_value str
      str.gsub('&quot;', '"') 
    end
    
    def office_string_escape str
      '"' + str.gsub('&quot;', '""') + '"'
    end
    
    # returns the parent table and removes template/examsple rows, also returns
    # the first template/example row (to find cell styles for instance)
    def retrieve_template_row doc, first_row_is_header

      root = doc.root

      count = 0
      template_row = nil

      table = root.elements['//table:table']
      table.each_element('table:table-row') do |row|

        if first_row_is_header && count == 0
          # keep the header row
        else

          table.delete_element(row)

          template_row = row unless template_row

        end

        count += 1

      end

      raise "Cannot find template row in #{@ods_template}" unless template_row

      return table, template_row

    end

    def create_temporary xml, to_filename

      to = Pathname.new(to_filename)

      temp_file = Tempfile.new('office-gen', to.dirname.to_s)

      File.open(temp_file, 'w') { |stream| stream.puts xml }

      temp_file

    end

    def write_result_content doc, content_xml_file, result_xml_file, first_row_is_header, &data_iterator

      file = File.new(content_xml_file)

      xml = REXML::Document.new(file)

      table, template_row = retrieve_template_row(xml, first_row_is_header)

      table.add_text "%%%Insert data here%%%\n"

      temp_file = create_temporary(xml, result_xml_file)

      File.open(result_xml_file, 'w') do |stream|

        @out_stream = stream

        File.open(temp_file, 'r').each do |line|

          if line =~ /(.*)%%%Insert data here%%%(.*)/
            @out_stream.write $1
            data_iterator.call(doc)
            @out_stream.write $2
          else
            @out_stream.write line
          end

        end

      end

    end

  end

end
