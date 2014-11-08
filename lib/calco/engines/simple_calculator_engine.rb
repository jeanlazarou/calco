require 'calco'

require_relative 'calculator_builtin_functions'

module Calco

  #
  # A simple engine computing the formulas.
  #
  # It does not cover some cases like
  #   * ahead formulas (formula using higher cells)
  #
  # Has no optimization, could make blow variable space...
  #
  # Is not thread safe if the same engine is used by several threads.
  # When 'save' is called an internal context is used and is the
  # weakness for thread safety...
  #
  class SimpleCalculatorEngine < Calco::DefaultEngine

    include CalculatorBuiltinFunctions
    
    def save doc, output, &data_iterator
    
      @context = Object.new.send(:binding)
      
      super
      
    end
    
    def write_row sheet, row_id
    
      return if sheet.has_titles? && row_id == 0
      
      names = column_names(sheet)
      
      cells = sheet.row(row_id)

      values, longest_name = compute_cell_values(row_id, cells, names)
      
      values.each_with_index do |value, i|
      
        next unless value
        
        value = value.to_s if value.is_a?(Date) || value.is_a?(Time)
        
        @out_stream.write "%#{longest_name}s = #{value.inspect}\n" % names[i]
        
      end
      
      @out_stream.write "\n"
      
    end
    
    def generate_cell row_number, column, cell_style, column_style, column_type
      
      return nil unless column

      if column.absolute_row
        return nil if column.absolute_row != row_number
      end
      
      if column.respond_to?(:value)
      
        if column.value.is_a?(Date)
          %{Date.parse("#{column.value}")}
        else
          column.value.inspect
        end
        
      else
        column.generate(row_number)
      end
      
    end

    def column_reference element, row
    
      name = super.downcase
      
      name.gsub!(%r{[$]([a-z]+)[$](\d+)}, '\1\2')

      name
      
    end
    
    private
    
    # returns an array of column title names, variable names or nil(s)
    def column_names sheet
    
      names = []
      
      titles = sheet.column_titles
      
      sheet.each_cell_definition do |column, i|
      
        if titles[i]

          names << titles[i]
          
        elsif column.respond_to?(:value_name)
          
          names << column.value_name
          
        else
          
          names << nil
          
        end
        
      end
      
      names
      
    end

    def compute_cell_values row_id, cells, names
    
      values = []
      longest_name = 0

      cells.each_with_index do |cell, i|

        name = "#{Calco::COLUMNS[i]}#{row_id}"
        
        if cell
    
          res = eval("#{name.downcase} = #{cell}", @context)
          
        else
          
          res = nil
          
        end
        
        values << res

        names[i] = name unless names[i]
        
        longest_name = [longest_name, names[i].length].max
        
      end
      
      return values, longest_name

    end
    
  end
  
end
