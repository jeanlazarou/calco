require_relative 'element'

module Calco

  class Constant < Element

    include Operators

    attr_accessor :value

    def initialize value
      @value = value
    end

    def generate row
      @engine.value(self)
    end

    def self.wrap value

      if value.respond_to?(:generate)
        value
      else
        Constant.new(value)
      end

    end

  end

end
