# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)
require "poi_xls_rails/version"

Gem::Specification.new do |s|
  s.name        = "poi_xls_rails"
  s.version     = PoiXlsRails::VERSION
  s.authors     = ["Enable Labs Pairing 3"]
  s.email       = ["pairing3@enablelabs.com"]
  s.homepage    = ""
  s.summary     = %q{TODO: Write a gem summary}
  s.description = %q{TODO: Write a gem description}

  s.rubyforge_project = "poi_xls_rails"

  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
end
