require 'csv'
require 'calco'

module Calco

  class CSVEngine < Calco::DefaultEngine

    def initialize col_sep = ',', quote_char = '"'
      @col_sep, @quote_char = col_sep, quote_char
    end

    def empty_row sheet
      @out_stream.write CSV.generate_line([])
    end

    def write_row sheet, row_id

      row_id += 1 if row_id > 0 && sheet.has_titles?

      cells = sheet.row(row_id)

      @out_stream.write CSV.generate_line(cells, :col_sep => @col_sep, :quote_char => @quote_char)

    end

    def generate_cell row_number, column, cell_style, column_style, column_type

      if column.is_a?(Calco::Formula)

        res = super

        if column_type =~ /^[$]/
          res = "=DOLLAR(#{res})"
        elsif column_type == '%'
          res = "=(#{res})%"
        else
          res = "=#{res}"
        end

      elsif column.respond_to?(:value)

        value = column.value

        style = cell_style ? "#{cell_style.generate(row_number)}" : ''

        if value.is_a?(Time)
          res = "=TIMEVALUE(\"#{value.strftime("%H:%M:%S")}\")#{style}"
        elsif value.is_a?(Date)
          res = "=DATEVALUE(\"#{value.strftime("%Y-%m-%d")}\")#{style}"
        elsif value.is_a?(Numeric)

          if column_type =~ /^[$]/
            res = "=DOLLAR(#{value}#{style})"
          elsif column_type == '%'
            res = "=(#{value}*100)%"
          else
            res = "#{value}#{style}"
          end

        else
          res = "#{value}#{style}"
        end

      end

      res

    end

    def style statement, row
      "+STYLE(#{statement.generate(row)})"
    end

    def current
      'CURRENT()'
    end

  end

end
