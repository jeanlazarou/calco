require_relative 'element'

module Calco

  class Or < Element

    def initialize condition1, condition2

      super()

      @condition1, @condition2 = Constant.wrap(condition1), Constant.wrap(condition2)

    end

    def generate row

      condition1 = generate_operand(@condition1, row)
      condition2 = generate_operand(@condition2, row)

      "OR(#{condition1};#{condition2})"

    end

  end

end
