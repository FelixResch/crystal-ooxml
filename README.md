# Crystal OOXML library

This library offers very basic support for reading and writing data from [OOXML]() files. For now only spreadsheets are supported, but support for other formats is planned as well.

## Installation

Add this to your application's `shard.yml`:

```yaml
dependencies:
  ooxml:
    github: FelixResch/ooxml
```

## Usage

Imagine a spreadsheet with data like this:

|   | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | Name | Birthday | Function | | |
| 2 | Alex | today | none | | |
| 3 | | | | | |

To read data from a worksheet use the following syntax:

```crystal
require "ooxml"

spreadsheet = OOXML::Spreadsheet.new(OOXML::Document.new("spreadsheet.xlsx"))

spreadsheet.worksheets.each do |worksheet|
    puts worksheet.name
end

users = spreadsheet.worksheets["Users"]

puts users["A2"] # => prints Alex

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
