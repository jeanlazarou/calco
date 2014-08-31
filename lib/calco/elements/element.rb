require_relative 'operator'

module Calco

  class Element

    attr_accessor :reference_type, :column, :absolute_row

    def initialize
      @absolute_row = nil
      @reference_type = :normal
    end

    def absolute_row= row_number
      @absolute_row = row_number
      @reference_type = :absolute
    end
    
    def generate_operand o, row
    
      if o.respond_to?(:as_operand)
        o.as_operand(row)
      else
        o.generate(row)
      end

    end

  end
  
end
