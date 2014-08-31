require_relative 'element'

module Calco

  class Formula < Element

    include Operators

    def initialize statement

      super()

      @statement = Constant.wrap(statement)

    end

    def [] range
      
      raise ArgumentError, "Expected Range got #{range.class}" unless range.is_a?(Range)
      raise ArgumentError, "Invalid start of range (must be > 0, was #{range.first})" unless range.first > 0
      
      Aggregator.new(self, range)
      
    end

    def generate row

      if self.reference_type == :absolute and self.absolute_row != row
        nil
      else
        @statement.generate(row)
      end

    end
    
    def as_operand row
      @engine.column_reference(self, row)
    end

  end

end
