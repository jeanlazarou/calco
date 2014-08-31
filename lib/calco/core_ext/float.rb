class Float

  alias_method '_calco_original_+', :+
  alias_method '_calco_original_-', :-
  alias_method '_calco_original_*', :*
  alias_method '_calco_original_/', :/
  
  %w[+ - / *].each do |op|

    define_method(op) do |arg|
    
      if arg.respond_to?(:generate)
        Calco::Operation.new(op, self, arg)
      else
        self.send("_calco_original_#{op}", arg)
      end
      
    end

  end
  
end
