require 'date'
require 'time'

require 'calco'

module Calco

  module CalculatorBuiltinFunctions
    
    def YEAR date
      date.year
    end
    
    def TODAY
      Date.today
    end
    
    def DATEVALUE text
      Date.parse(text)
    end
    
    def TIMEVALUE text
      Time.parse(text)
    end
    
    def LEFT text, n
      text[0, n]
    end
    
  end

end
