require 'calco'

module Calco

  class DefaultEngine

    # output can be a String (as a file name) or an output stream
    # examples:
    #     engine.save doc, "my_doc.txt"  do |spreadsheet|
    #        ...
    #     end
    #
    #     engine.save doc, $stdout  do |spreadsheet|
    #        ...
    #     end
    def save doc, output, &data_iterator
    
      if output.respond_to?(:write)
        @out_stream = output
      else
        @out_stream = open(output, "w")
      end
    
      data_iterator.call(doc)
      
    end

    def empty_row sheet
    end

    def write_row sheet, row_id
    
      cells = sheet.row(row_id)
      
      cells.each_index do |i|

        cell = cells[i]

        @out_stream.write "#{COLUMNS[i]}#{row_id}: #{cell}\n"

      end
      
      @out_stream.write "\n"
      
    end

    def generate row_number, columns, cell_styles, column_styles, column_types

      cells = []

      columns.each_with_index do |column, i|
        cells << generate_cell(row_number, column, cell_styles[i], column_styles[i], column_types[i])
      end

      cells

    end

    def generate_cell row_number, column, cell_style, column_style, column_type

      cell = '' unless column
      cell = column.generate(row_number) if column

      if cell_style
        cell = cell.to_s + cell_style.generate(row_number)
      end

      cell

    end

    def reference_types
      {
          :current => proc { |element, r, c| current },
          :normal => proc { |element, r, c| "#{c}#{r}" },
          :absolute => proc { |element, r, c| "$#{c}$#{element.absolute_row}" },
          :absolute_row => proc { |element, r, c| "#{c}$#{element.absolute_row}" },
          :absolute_column => proc { |element, r, c| "$#{c}#{r}" }
      }
    end

    def column_reference element, row
      reference_types[element.reference_type].call(element, row, element.column)
    end

    def range_reference element, row, range
    
      from = reference_types[element.reference_type].call(element, range.first, element.column)
      
      if range.grouping_range?
        to = row - 1
      elsif range.last == -1
        to = DefaultEngine.row_infinity
      else
        to = range.last
      end
      
      to = reference_types[element.reference_type].call(element, to, element.column)
      
      "#{from}:#{to}"
      
    end

    def value element
      if ! element.respond_to?(:value)
        element.inspect
      elsif element.value.is_a?(Time)
        element.value.strftime("%H:%M:%S")
      elsif element.value.is_a?(Date)
        element.value.strftime("%Y-%m-%d")
      else
        element.value.inspect
      end
    end

    def style statement, row
      "; apply_style(#{statement.generate(row)})"
    end

    def operator op
      op
    end

    def current
      'self'
    end

    def self.row_infinity
      @@row_infinity
    end
    
    def self.row_infinity= value
      @@row_infinity = value
    end

    @@row_infinity = 1048576

  end

end
