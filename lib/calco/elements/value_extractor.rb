require_relative 'element'

module Calco

  class ValueExtractor < Element

    def initialize variable, reference_type
      @variable, @reference_type = variable, reference_type
      @variable.reference_type = reference_type
    end

    def column= column_name
      @variable.column = column_name
    end

    def value_name
      @variable.name
    end

    def value
      @variable.value
    end

    def absolute_row
      @variable.absolute_row
    end

    def reference_type
      @variable.reference_type
    end

    def generate row
      
      return nil if @variable.absolute_row && @variable.absolute_row != row

      @engine.value(self)
      
    end

  end

end
