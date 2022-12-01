# spreadsheet

[![Latest Version on Packagist][ico-version]][link-packagist]
[![Software License][ico-license]](LICENSE.md)

Simple spreadsheet reader supporting XLS, XLSX and CSV files. 

## Install

Via Composer

``` bash
$ composer require vakata/spreadsheet
```

## Usage

``` php
foreach (\vakata\spreadsheet\Reader::fromFile('Book1.xls') as $k => $row) {
    var_dump($row);
}
foreach (\vakata\spreadsheet\Reader::fromFile('Book1.xlsx') as $k => $row) {
    var_dump($row);
}
foreach (\vakata\spreadsheet\Reader::fromFile('Book1.csv') as $k => $row) {
    var_dump($row);
}
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

