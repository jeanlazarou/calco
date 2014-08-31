# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'calco/version'

Gem::Specification.new do |gem|
  gem.name          = "calco"
  gem.version       = Calco::VERSION
  gem.authors       = ["Jean Lazarou"]
  gem.email         = ["jean.lazarou@alef1.org"]
  gem.description   = %q{Implements a DSL used to create and define the content of spreadsheet documents}
  gem.summary       = %q{DSL for spreadsheet documents}
  gem.homepage      = "https://github.com/jeanlazarou/calco"

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.require_paths = ["lib"]
  
  gem.add_development_dependency('rspec')
  
  gem.add_dependency 'rubyzip', '~> 1.1.0'
end
