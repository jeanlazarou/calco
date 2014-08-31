module Calco

  module Operators

    %w[+ - / * == != < > >= <=].each do |op|

      op_str = op == '==' ? '=' : op

      define_method(op) do |arg|
        Operation.new(op_str, self, arg)
      end

    end

  end

end
