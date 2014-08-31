require_relative 'element'

module Calco

  class If < Element

    def initialize condition, _then, _else

      super()

      @condition, @_then, @_else = condition, Constant.wrap(_then), Constant.wrap(_else)

    end

    def generate row

      _then = generate_operand(@_then, row)
      _else = generate_operand(@_else, row)

      "IF(#{@condition.generate(row)};#{_then};#{_else})"

    end

  end

end
