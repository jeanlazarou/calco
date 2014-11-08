require 'set'

require_relative 'style'
require_relative 'definition_dsl'

module Calco

  letters = (65...65 + 26).collect { |i| i.chr }
  COLUMNS = letters + letters.collect { |letter| (65...65 + 26).collect { |i| letter + i.chr } }.flatten

  # don't use method names, of the Sheet class, as variable or formula or even 'sheet_name'
  class Sheet

    include BuilderDSL
    
    attr_accessor :definitions
    attr_accessor :sheet_name, :column_titles

    def initialize sheet_name = 'Sheet', engine = DefaultEngine.new

      @sheet_name = sheet_name

      @columns = []
      @cell_styles = []
      @column_types = []
      @title_styles = []
      @column_styles = []
      @assigned_column_index = 0

      @columns_by_id = {}
      
      @save_contents = []
      
      @has_titles = false
      @column_titles = []

      @engine = engine

    end

    def owner= owner
      @owner = owner
    end

    def current
      Current.new
    end

    def column_style index
      @column_styles[index]
    end

    def title_style index
      @title_styles[index]
    end

    # name is a numeric or string constant, a variable name, a ValueExtractor, a formula name, empty (skip)
    # options is a has than can contain next keys:
    #    style        => a dynamic (or computed) style for the cell
    #    column_style => style name for the column
    #    title_style  => style name for the column title (if row 0 contains the title)
    #    type         => type for the column values
    #    title        => the title of the column, supposed to appear at first row
    #    id           => id associated to the column
    def column name, options = {}

      if name.is_a?(Numeric) || name.is_a?(String)
        @columns << Constant.wrap(name)
      elsif name.is_a?(ValueExtractor)
        @columns << name
        name.column = COLUMNS[@assigned_column_index]
      elsif name == Empty.instance
        @columns << nil
      elsif @definitions.variable?(name)
        @definitions.variable(name).column = COLUMNS[@assigned_column_index]
        @columns << @definitions.variable(name)
      elsif @definitions.formula?(name)
        @definitions.formula(name).column = COLUMNS[@assigned_column_index]
        @columns << @definitions.formula(name)
      else
        raise "Unknown function or variable '#{name}'"
      end

      raise ArgumentError, "Options should be a Hash" unless options.is_a?(Hash)

      @columns_by_id[options[:id]] = @assigned_column_index if options[:id]
      
      title = nil
      title = options[:title] if options[:title]

      style = nil
      style = Style.new(options[:style]) if options[:style]

      type = nil
      type = options[:type] if options[:type]

      column_style = nil
      column_style = options[:column_style] if options[:column_style]

      title_style = nil
      title_style = options[:title_style] if options[:title_style]

      @column_titles << title
      @has_titles = @has_titles || !title.nil?

      @cell_styles << style
      @column_types << type
      @title_styles << title_style
      @column_styles << column_style

      @assigned_column_index += 1

    end

    def replace_and_clear replacements

      changed = Set.new
      
      replacements.each do |id, new_content|
        
        replace_content id, new_content
        
        changed << @columns_by_id[id]
        
      end
      
      @columns.each_with_index do |current_content, index|
        
        next if changed.include?(index)
        
        @save_contents[index] = current_content
        @columns[index] = nil
      
      end
      
    end
    
    def replace_content column_id, new_content
    
      raise "Column id '#{column_id}' not found" unless @columns_by_id.include?(column_id)
      
      index = @columns_by_id[column_id]
      
      current_content = @columns[index]
      
      @save_contents[index] = current_content

      if new_content.nil?
        @columns[index] = nil
        return
      end
      
      if new_content.is_a?(Numeric) || new_content.is_a?(String)
        new_content = Constant.new(new_content)
        assign_engine new_content, @engine
      elsif @definitions.variable?(new_content)
        new_content = ValueExtractor.new(@definitions.variable(new_content), :normal)
        assign_engine new_content, @engine
      elsif @definitions.formula?(new_content)
        new_content = @definitions.formula(new_content)
      else
        raise "Unknown function or variable '#{new_content}' for replacement"
      end
      
      new_content.column = current_content.column
      
      @columns[index] = new_content
      
    end
    
    def restore_content
    
      @save_contents.each_with_index do |saved_content, index|
        @columns[index] = saved_content if saved_content
      end
      
      @save_contents.clear
      
    end
    
    def skip
      Empty.instance
    end

    # type can be :normal, :absolute, :absolute_row, :absolute_column
    def value_of name, type = :normal

      if @definitions.variable?(name)
        ValueExtractor.new(@definitions.variable(name), type)
      else
        raise "Unknown variable #{name}"
      end

    end

    def [] name

      return @definitions.variable(name).value if @definitions.variable?(name)

      nil

    end

    def []= variable_name, new_value

      raise "Unknown variable '#{variable_name}'" unless @definitions.variable?(variable_name)

      @definitions.variable(variable_name).value = new_value

    end

    # record is a Hash or any object implementing #each
    # with a block expecting two parameters:
    #    * the key used (the variable name)
    #    * the value to assign
    def record_assign record

      record.each do |name, value|
        self[name] = value
      end

    end

    def has_titles?
      @has_titles
    end

    def has_titles flag
      @has_titles = @has_titles || flag
    end

    # Returns an array of String objects generates by the engine
    # attached to this Sheet's Spreadsheet.
    def row row_number

      if row_number == 0

        cells = []

        if @has_titles

          @column_titles.each do |title|
            cells << title ? title : ''
          end

        end

        cells
        
      else

        @engine.generate(row_number, @columns, @cell_styles, @column_styles, @column_types)

      end

    end

    # set "next row" as empty, delegated to current engine
    def empty_row
      @engine.empty_row
    end
    
    # Calls the passed block for every Element, the cell definition,
    # attached to columns for this Sheet.
    #
    # Depending an the block's arity, passes the Element or the Element
    # and the index to the given block.
    #
    # Examples:
    #   each_cell_definition {|column_def| ... }
    #   each_cell_definition {|column_def, index| ... }
    def each_cell_definition &block
    
      if block.arity == 1
        
        @columns.each do |column|
          yield column
        end
        
      else
        
        @columns.each_with_index do |column, i|
          yield column, i
        end
        
      end
      
    end

    # data is a hash passed to #record_assign
    def write_row row_id, data = nil
    
      record_assign data if data
    
      @engine.write_row self, row_id
      
    end
    
    def compile engine = @engine
    
      @engine = engine
      
      @definitions.formulas.each_value do |formula|
        assign_engine formula, engine
      end
      @definitions.variables.each_value do |variable|
        assign_engine variable, engine
      end
    
      @columns.each do |column|
        assign_engine column, engine
      end
    
      @cell_styles.each do |style|
        assign_engine style, engine
      end
      
    end
    
    def method_missing sym, *args
      # to avoid some strange behavior when some spec fails (ending with a 
      # call to #to_ary) enable next line to help fixing the problem
      #super.respond_to?(sym)
      @definitions.resolve! sym, *args
    end

    private
    
    def assign_engine el, engine

      return unless el.is_a?(Calco::Element) || el.is_a?(Calco::Style)

      current_engine = el.instance_variable_get(:@engine)
      
      unless current_engine == engine
        
        el.instance_variable_set(:@engine, engine)

        el.instance_variables.each do |var|
          
          val = el.instance_variable_get(var)
          
          if val.respond_to?(:each)
          
            val.each do |item|
              assign_engine item, engine
            end

          else
          
            assign_engine val, engine
          
          end
          
        end
        
      end
      
    end
    
  end

end
