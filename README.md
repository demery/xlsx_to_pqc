# XlsxToMarmite

Convert PQC structural metadata spreadsheet and directory contents into
PQC structural metadata. 

```xml
<record>
    <bib_id>9963173193503681</bib_id>
    <pages>
        <page number="1" seq="1" id="0001" image.defaultscale="3" side="recto" 
          image.id="0001" image="0001" visiblepage="1r" display="true">
          <tocentry name="toc">Pio, Alberto (1512-1518)</tocentry>
        </page>
        <page number="2" seq="2" id="0002" image.defaultscale="3" side="verso" 
          image.id="0002" image="0002" visiblepage="1v" display="true"/>
        <page number="3" seq="3" id="0003" image.defaultscale="3" side="recto" 
          image.id="0003" image="0003" visiblepage="2r" display="true"/>
        <page number="4" seq="4" id="0004" image.defaultscale="3" side="verso" 
          image.id="0004" image="0004" visiblepage="2v" display="true">
          <tocentry name="toc">Table, f. 2v [=3v]</tocentry>
        </page>
        <page number="5" seq="5" id="0005" image.defaultscale="3" side="recto" 
          image.id="0005" image="0005" visiblepage="3r" display="true"/>
        <page number="6" seq="6" id="0006" image.defaultscale="3" side="verso" 
          image.id="0006" image="0006" visiblepage="3v-4r" display="true">
          <tocentry name="ill">Decorated initial, Initial P, p. 3</tocentry>
          <tocentry name="ill">Foliate design, p. 3</tocentry>
        </page>
        <page number="7" seq="7" id="0007" image.defaultscale="3" side="recto" 
          image.id="0007" image="0007" visiblepage="4v" display="true"/>
        <page number="8" seq="8" id="0002a" image.defaultscale="3" 
          image.id="0002a" image="0002a" display="false"/>          
        <page number="9" seq="9" id="0002b" image.defaultscale="3" 
          image.id="0002b" image="0002b" display="false"/>
        <page number="10" seq="10" id="reference" image.defaultscale="3" 
          image.id="reference" image="reference" display="false"/>
    </pages>
</record>
``` 

TODO: What if the page is "3v-4r"? Is that recto or verso?

> It looks like the DLA and Marmite always alternate, recto/verso, regardless;
> see:
>
> <http://mdproc.library.upenn.edu:9292/records/9959371403503681/create?format=structural>

TODO: Does a non-display image have `seq`or `visiblepage` values?

PQC structural data will be in an XLSX spreadsheet named `pqc_structural.xlsx`
and formatted thus:

| ARK ID                | PAGE SEQUENCE | VISIBLE PAGE           | TOC ENTRY                | ILL ENTRY                                                | FILENAME | NOTES  |
|-----------------------|---------------|------------------------|--------------------------|----------------------------------------------------------|----------|--------|
| ark:/99999/fk42244n9f | 1             | 1r                     | Pio, Alberto (1512-1518) |                                                          | 0001.tif |        |
| ark:/99999/fk42244n9f | 2             | 1v                     |                          |                                                          | 0002.tif |        |
| ark:/99999/fk42244n9f | 3             | 2r                     |                          |                                                          | 0003.tif |        |
| ark:/99999/fk42244n9f | 4             | 2v                     | Table, f. 2v [=3v]       |                                                          | 0004.tif |        |
| ark:/99999/fk42244n9f | 5             | 3r                     |                          |                                                          | 0005.tif |        |
| ark:/99999/fk42244n9f | 6             | 3v-4r                  |                          | Decorated initial, Initial P, p. 3\|Foliate design, p. 3 | 0006.tif |        |
| ark:/99999/fk42244n9f | 7             | 4v                     |                          |                                                          | 0007.tif |        |


The spreadsheet will be in the same directory as the images, which will have
this structure:

```text
    ark+=99999=fk45t4vg3q
    ├── 0001.tif
    ├── 0002.tif
    ├── 0002a.tif
    ├── 0002b.tif
    ├── 0003.tif
    ├── 0004.tif
    ├── 0005.tif
    ├── 0006.tif
    ├── pqc_descriptive.xlsx
    ├── pqc_structural.xlsx
    └── reference.tif
```

Notice that the directory contains three images that a present in the structural
XML but not listed in the XLSX: `0002a.tif`, `0002b.tif`, and `reference.tif`.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'fixtures'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install xlsx_to_marmite

## Usage

TODO: Write usage instructions here

