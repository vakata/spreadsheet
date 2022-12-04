# spreadsheet

[![Latest Version on Packagist][ico-version]][link-packagist]
[![Software License][ico-license]](LICENSE.md)

Simple spreadsheet reader/writer supporting XLS (read-only), XLSX, XML and CSV files. 
The classes try not to load all the data at once so that fairly large files are supported using iterators.

## Install

Via Composer

``` bash
$ composer require vakata/spreadsheet
```

## Usage

``` php
// you can also write to browser or to stream (additional options are available for each format)
foreach (\vakata\spreadsheet\Writer::toFile('test.xlsx')->fromArray([
    [1,"asdf","2022-02-10"],
    [2,"test","2010-11-10"]
]);
// you can also read from stream
foreach (\vakata\spreadsheet\Reader::fromFile('test.xlsx') as $k => $row) {
    var_dump($row);
}
// or
var_dump(\vakata\spreadsheet\Reader::fromFile('test.xlsx')->toArray());
```

## Contributing

Please see [CONTRIBUTING](CONTRIBUTING.md) for details.

## Security

If you discover any security related issues, please email github@vakata.com instead of using the issue tracker.

## Credits

- [kterziev][link-mainauthor]
- [vakata][link-author]

## License

The MIT License (MIT). Please see [License File](LICENSE.md) for more information.

[ico-version]: https://img.shields.io/packagist/v/vakata/spreadsheet.svg?style=flat-square
[ico-license]: https://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square
[ico-downloads]: https://img.shields.io/packagist/dt/vakata/spreadsheet.svg?style=flat-square

[link-packagist]: https://packagist.org/packages/vakata/spreadsheet
[link-downloads]: https://packagist.org/packages/vakata/spreadsheet
[link-mainauthor]: https://github.com/kterziev
[link-author]: https://github.com/vakata

