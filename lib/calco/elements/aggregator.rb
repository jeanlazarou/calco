require_relative 'element'

module Calco

  class Aggregator < Element

    def initialize variable, range
      @variable, @range = variable, range
    end

    def generate row
      @engine.range_reference(@variable, row, @range)
    end

  end

end
