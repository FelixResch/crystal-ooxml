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

workbook = OOXML::Workbook.new(OOXML::Document.new("spreadsheet.xlsx"))

workbook.worksheets.each do |worksheet|
    puts worksheet.name
end

users = workbook.worksheets["Users"]

puts users["A2"] # => prints Alex
```

## Roadmap

For an initial version these things still need to be done:

- [ ] Writing of Workbook files
- [ ] Unified table-stuctured data interface for crystal
- [ ] Plugin System for function evaluation
- [ ] Documentation and Optimization

Currently the following points are in development or are already finished:

- [x] Reading of Workbook files


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
