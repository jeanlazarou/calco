require 'date'
require 'time'

require_relative 'elements/if'
require_relative 'elements/or'
require_relative 'elements/empty'
require_relative 'elements/constant'
require_relative 'elements/current'
require_relative 'elements/element'
require_relative 'elements/formula'
require_relative 'elements/variable'
require_relative 'elements/operation'
require_relative 'elements/value_extractor'
require_relative 'elements/builtin_function'

require_relative 'core_ext/float'
require_relative 'core_ext/range'
require_relative 'core_ext/fixnum'
require_relative 'core_ext/string'

module Calco

  module DefinitionDSL
  
    def set variable_assign

      name = variable_assign.first[0]
      value = variable_assign.first[1]

      raise "Variable '#{name}' already set" if @variables.include?(name)

      if value =~ /^\d{1,2}:\d{1,2}(:\d{1,2})?$/
        value = Time.parse(value)
      elsif value =~ /^\d{1,4}[-\/]\d{1,2}[-\/]\d{1,2}$/
        value = Date.parse(value)
      end
      
      @variables[name] = Variable.new(name, value)

    end

    def function formula_definition

      name = formula_definition.first[0]
      value = formula_definition.first[1]

      raise "Function '#{name}' already defined" if @formulas.include?(name)
      
      @formulas[name] = Formula.new(value)

    end

  end
  
  module BuilderDSL
  
    def _if condition, then_, else_
      If.new(condition, then_, else_)
    end

    def _or condition1, condition2
      Or.new(condition1, condition2)
    end
    
    def builtin_function? sym
      BuiltinFunction.exists_function?(sym)
    end
    
    def builtin_function sym, args
      BuiltinFunction.create_function(sym, args)
    end
    
  end
  
  class Definitions
  
    include BuilderDSL
    include DefinitionDSL

    attr_reader :variables, :formulas
    
    def initialize spreadsheet
      
      @formulas = {}
      @variables = {}
      
      @spreadsheet = spreadsheet
      
    end
  
    def resolve! sym, *args

      if @variables.include?(sym)
        @variables[sym]
      elsif @formulas.include?(sym)
        @formulas[sym]
      elsif builtin_function?(sym)
        builtin_function(sym, args)
      else
        raise "Unknown function or variable '#{sym}'"
      end

    end
    
    def formula? name
      @formulas.include?(name)
    end
    
    def formula name
      @formulas[name]
    end
    
    def variable? name
      @variables.include?(name)
    end
    
    def variable name
      @variables[name]
    end
    
    def method_missing sym, *args
      resolve! sym, *args
    end

  end
  
end
