require "calco/version"

require "calco/spreadsheet"
require "calco/time_functions"
require "calco/date_functions"
require "calco/math_functions"
require "calco/string_functions"

def spreadsheet engine = Calco::DefaultEngine.new, &block

  document = Calco::Spreadsheet.new(engine)

  document.instance_eval(&block)

  document

end
