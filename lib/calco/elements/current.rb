require_relative 'element'

module Calco

  class Current < Element

    include Operators

    def initialize
      @reference_type = :current
    end

    def generate row
      @engine.current
    end

  end

end
