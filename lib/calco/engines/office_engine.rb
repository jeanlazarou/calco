require 'date'

require 'calco'
require 'calco/xml_builder'

require_relative 'office_file_manager'

module Calco

  class OfficeEngine < DefaultEngine

    def initialize ods_template
      @ods_template = ods_template
    end
    
    # output is a String (as a file name)
    def save doc, to_filename, &data_iterator

      @file_manager = OfficeFileManager.new(@ods_template)
      
      data_iterator.call(doc)

      @file_manager.save doc, to_filename
      
    end

    def empty_row sheet
      @file_manager.add_empty_row sheet
    end
    
    def write_row sheet, row_id

      return if row_id == 0 && sheet.has_titles?
      
      row_id += 1 # office sheet indexes start at 1
      
      cells = sheet.row(row_id)

      @file_manager.add_row sheet do |stream|
      
        cells.each_index do |i|

          cell = cells[i]

          stream.write cell

        end

      end
      
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

  end

end
