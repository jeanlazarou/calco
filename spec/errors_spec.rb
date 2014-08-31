require 'calco'

RSpec.configure do |c|
  c.alias_example_to :detect
end

describe 'Spreadsheet errors' do

  detect "variable does not exist" do

    expect {

      spreadsheet do

        sheet do

          column value_of(:some_var)

        end

      end

    }.to raise_error("Unknown variable some_var")

  end

  detect "function using an unknown variable" do

    expect {

      doc = spreadsheet do

        definitions do

          function some_var + 1

        end

      end

    }.to raise_error("Unknown function or variable 'some_var'")

  end

  detect "assiging value to unknown variable" do

    doc = spreadsheet do

      sheet do

      end

    end

    expect {

      doc.save($stdout) do |spreadsheet|

        sheet = spreadsheet.current

        sheet[:some_var] = 12

      end

    }.to raise_error("Unknown variable 'some_var'")

  end

  detect "reference to an unknown function" do

    expect {

      spreadsheet do

        sheet do

          column :some_function

        end

      end

    }.to raise_error("Unknown function or variable 'some_function'")

  end

  detect "using an unknown function" do

    expect {

      spreadsheet do

        definitions do

          function name: some_function

        end

      end

    }.to raise_error("Unknown function or variable 'some_function'")

  end

  detect "declaring a variable twice" do

    expect {

      spreadsheet do

        definitions do

          set var: 12
          set var: "hello"

        end

      end

    }.to raise_error("Variable 'var' already set")

  end

  detect "declaring a function twice" do

    expect {

      spreadsheet do

        definitions do

          function f: 12
          function f: "hello"

        end

      end

    }.to raise_error("Function 'f' already defined")

  end

  detect "unnamed option for column" do

    expect {

      spreadsheet do

        definitions do

          set price: 11

        end

        sheet do
          column :price, 'pp'
        end

      end

    }.to raise_error(ArgumentError, "Options should be a Hash")

  end

  detect "unknown column id" do

    expect {

      doc = spreadsheet do

        definitions do

          set a: 'a'

        end

        sheet do
          column :a
        end

      end

      doc.current.replace_content :no_id, nil

    }.to raise_error(RuntimeError, "Column id 'no_id' not found")

  end

end
