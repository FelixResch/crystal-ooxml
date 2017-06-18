# ooxml

TODO: Write a description here

## Installation

Add this to your application's `shard.yml`:

```yaml
dependencies:
  ooxml:
    github: FelixResch/ooxml
```

## Usage

```crystal
require "ooxml"
```

```crystal
spreadsheet = OOXML::Spreadsheet.new(OOXML::Document.new("spreadsheet.xlsx"))

spreadsheet.worksheets.each do |worksheet|
    puts worksheet.name
end
```

## Development

TODO: Write development instructions here

## Contributing

1. Fork it ( https://github.com/FelixResch/ooxml/fork )
2. Create your feature branch (git checkout -b my-new-feature)
3. Commit your changes (git commit -am 'Add some feature')
4. Push to the branch (git push origin my-new-feature)
5. Create a new Pull Request

## Contributors

- [FelixResch](https://github.com/FelixResch) felix - creator, maintainer
