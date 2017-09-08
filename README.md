# spreadsheet

[![Latest Version on Packagist][ico-version]][link-packagist]
[![Software License][ico-license]](LICENSE.md)
[![Build Status][ico-travis]][link-travis]
[![Scrutinizer Code Quality][ico-code-quality]][link-scrutinizer]
[![Code Coverage][ico-scrutinizer]][link-scrutinizer]

Simple spreadsheet reader supporting XLS, XLSX and CSV files

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

## Testing

``` bash
$ composer test
```


## Contributing

Please see [CONTRIBUTING](CONTRIBUTING.md) for details.

## Security

If you discover any security related issues, please email github@vakata.com instead of using the issue tracker.

## Credits

- [vakata][link-author]
- [All Contributors][link-contributors]

## License

The MIT License (MIT). Please see [License File](LICENSE.md) for more information.

[ico-version]: https://img.shields.io/packagist/v/vakata/spreadsheet.svg?style=flat-square
[ico-license]: https://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square
[ico-travis]: https://img.shields.io/travis/vakata/spreadsheet/master.svg?style=flat-square
[ico-scrutinizer]: https://img.shields.io/scrutinizer/coverage/g/vakata/spreadsheet.svg?style=flat-square
[ico-code-quality]: https://img.shields.io/scrutinizer/g/vakata/spreadsheet.svg?style=flat-square
[ico-downloads]: https://img.shields.io/packagist/dt/vakata/spreadsheet.svg?style=flat-square
[ico-cc]: https://img.shields.io/codeclimate/github/vakata/spreadsheet.svg?style=flat-square
[ico-cc-coverage]: https://img.shields.io/codeclimate/coverage/github/vakata/spreadsheet.svg?style=flat-square

[link-packagist]: https://packagist.org/packages/vakata/spreadsheet
[link-travis]: https://travis-ci.org/vakata/spreadsheet
[link-scrutinizer]: https://scrutinizer-ci.com/g/vakata/spreadsheet
[link-code-quality]: https://scrutinizer-ci.com/g/vakata/spreadsheet
[link-downloads]: https://packagist.org/packages/vakata/spreadsheet
[link-author]: https://github.com/vakata
[link-contributors]: ../../contributors
[link-cc]: https://codeclimate.com/github/vakata/spreadsheet

