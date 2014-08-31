require_relative 'element'
require_relative 'aggregator'

module Calco

  class Variable < Element

    include Operators

    attr_accessor :value, :name

    def initialize name, value

      super()

      @name, @value = name, value

    end

    def [] range
      
      raise ArgumentError, "Expected Range got #{range.class}" unless range.is_a?(Range)
      raise ArgumentError, "Invalid start of range (must be > 0, was #{range.first})" unless range.first > 0
      
      Aggregator.new(self, range)
      
    end
    
    def generate row
      @engine.column_reference(self, row)
    end

  end

end
