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

// when writing XLSX files formatting options are available
$writer = \vakata\spreadsheet\Writer::toFile('test.xlsx', 'xlsx');
$driver = $writer->getDriver();
$driver->addSheet('Sheet name');
$driver->addHeaderRow(['', ['', null, 'LTBRD'], ['Names', null, null, null, 'c', null, 3], '', '', ''], false, false);
$driver->addHeaderRow(['', '№', 'Given', 'Surname','Family', 'Year']);
$driver->addRow([['group 1', null, null, null, 'CM', '999999', 1, 3 ], 1, 'Leopold', 'Sarah', 'Johnson', 1981], 'b');
$driver->addRow(['', 2, 'Phil', 'Stuart', 'Davidson', 1984], '', 'LTBR', '009900');
$driver->addRow(['', 3, 'Anne', 'Marie', 'Gordon', [1992, 'biu', null, null, null, '00FF00']]);
$driver->addRow([['group 2', null, null, null, 'CM', '999999', 1, 2 ], 4, 'George', '', 'Black', 1978]);
$driver->addRow(['', 5, 'David', '', 'Green', 1989]);
$driver->close();
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

