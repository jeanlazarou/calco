require 'singleton'

require_relative 'sheet'
require_relative 'definition_dsl'
require_relative 'engines/default_engine'

module Calco

  class Spreadsheet

    def initialize engine = DefaultEngine.new
      @engine = engine
      @sheets = []
      @sheets_by_name = {}
      @named_styles = {}
      @definitions = Definitions.new(self)
    end

    def save to_filename, &data_iterator
      @engine.save self, to_filename, &data_iterator
    end

    def definitions &block

      @block_def = block
      @definitions.instance_eval(&block)

    end

    def sheet name = nil, &block

      if @sheets_by_name[name]
        @current = sheet = @sheets_by_name[name]
      else

        sheet = create_sheet(name)

        @sheets << sheet

        name = "Sheet #{@sheets.size}" unless name

        @sheets_by_name[name] = sheet

      end

      sheet.instance_eval(&block) if block_given?
      sheet.compile

      self

    end

    def percentage_style name, decimal_places, min_digits, text
      @named_styles[name] = {type: :percentage, decimal_places: decimal_places, min_digits: min_digits, text: text}
    end

    def cell_style name, other_styles = {}

      other_styles[:type] = :cell

      @named_styles[name] = other_styles

    end

    def column_style name, width
      @named_styles[name] = {type: :column, width: width}
    end

    def row_style name, width
      @named_styles[name] = {type: :row, height: width}
    end

    def table_style name, styles = {}

      @named_styles[:type] = :table

      @named_styles[name] = styles

    end

    def row row_number
      @current.row row_number
    end

    def current name_or_index = nil

      if name_or_index
        @current = find_sheet!(name_or_index)
      elsif @current
        @current
      else
        sheet
        @current
      end

    end

    def current? name_or_index
      @current == find_sheet(name_or_index)
    end

    def [] name_or_index
      find_sheet(name_or_index)
    end

    def find_sheet name_or_index

      if name_or_index.is_a?(Fixnum)
        @sheets[name_or_index]
      else
        @sheets_by_name[name_or_index]
      end

    end

    def find_sheet! name_or_index

      sheet = find_sheet(name_or_index)

      raise "Unknown sheet '#{name_or_index}'" unless sheet

      sheet

    end

    def row_to_hash row_number

      hash = {}

      cells = row(row_number)

      cells.each_index do |i|

        cell = cells[i]

        hash[COLUMNS[i]] = cell

      end

      hash

    end

    def engine= new_engine

      @engine = new_engine

      @sheets.each do |sheet|
        sheet.compile @engine
      end

    end

    private

    def create_sheet name

        definitions = Definitions.new(self)

        definitions.instance_eval(&@block_def) if @block_def

        @current = Sheet.new(name, @engine)
        @current.definitions = definitions
        @current.owner = self

        @current

    end

  end

end
