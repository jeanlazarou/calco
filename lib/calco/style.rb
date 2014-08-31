module Calco

  class Style

    def initialize statement
      @statement = statement
    end

    def generate row
      @engine.style(@statement, row)
    end

  end

end
