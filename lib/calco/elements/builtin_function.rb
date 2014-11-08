require_relative 'element'

module Calco

  class BuiltinFunction < Element

    @@builtin_functions = {}

    # arity is an integer or :n for a variable number of arguments
    # type is the return type and expected to be a class object or an array
    # of class objects
    def self.declare name, arity, type = :any

      unless arity == :n || arity.is_a?(Integer)
        raise ArgumentError, "Artity must be an integer or :n but was a #{arity.class}"
      end
      unless type.is_a?(Class) || (type.respond_to?(:all?) && type.all?{|t| t.is_a?(Class)})
        raise ArgumentError, "Type should be a Class or an array of Class objects" 
      end
      
      @@builtin_functions[name.to_s.downcase] = [arity, type]

    end

    def self.undeclare name
      @@builtin_functions.delete(name.to_s.downcase)
    end

    def self.exists_function? name
      @@builtin_functions.include?(name.to_s.downcase)
    end

    def self.create_function name, args

      raise NameError, "Builtin function not found '#{name}'" unless exists_function?(name)

      name = name.to_s.upcase

      definition = @@builtin_functions[name.to_s.downcase]

      self.new(name, definition, args)

    end

    include Operators

    def initialize name, definition, args

      @arity = definition[0]
      @return_type = definition[1]

      unless @arity == :n

        if @arity != args.length
          raise ArgumentError, "Function #{name} requires #{@arity}, was #{args.length} (#{args.inspect})"
        end

      end

      @name, @args = name, args

    end

    def generate row

      sep = ''
      arguments = ''

      @args.each do |arg|

        arguments << sep
        arguments << (arg.respond_to?(:generate) ? arg.generate(row) : @engine.value(arg))

        sep = ', '

      end

      "#{@name}(#{arguments})"

    end

  end

end
