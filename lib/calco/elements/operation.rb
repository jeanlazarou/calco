require_relative 'element'

module Calco

  class Operation < Element

    include Operators

    def initialize operator, operand1, operand2
      @operator, @operand1, @operand2 = operator, Constant.wrap(operand1), Constant.wrap(operand2)
    end

    def generate row

      return "#{@engine.operator(@operator)}#{generate_operand(@operand1, row)}" unless @operand2
      
      operand1 = generate_operand(@operand1, row)
      operand2 = generate_operand(@operand2, row)

      "#{operand1}#{@engine.operator(@operator)}#{operand2}"

    end

    def as_operand row
      "(#{generate(row)})"
    end
    
    def -@
      Operation.new('-', self, nil)
    end
    
  end

end
