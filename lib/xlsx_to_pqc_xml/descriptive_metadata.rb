require 'nokogiri'

module XlsxToPqcXml
  ##
  # Class to generate PQC descriptive metadata XML for directory containing
  # images and a PQC descriptive metadata spreadsheet, using a hash spreadsheet
  # configuration.
  #
  # This class validates the spreadsheet based on the configuration.
  #
  # === Usage
  #
  # Given the following YAML configuration and a directory containing a
  # spreadsheet named +pqc_descriptive.xlsx+:
  #
  # Spreadsheet configuration +descriptive_config.yml+:
  #
  #     ---
  #     :sheet_name: 'Descriptive'
  #     :sheet_position: 0
  #     :heading_type: 'row'
  #     :attributes:
  #     - :attr: object_type
  #       :xml_element: type
  #       :headings:
  #       - OBJECT TYPE
  #       :requirement: required
  #     - :attr: unique_identifier
  #       :xml_element: ark
  #       :data_type: :ark
  #       :headings:
  #       - UNIQUE IDENTIFIER
  #       :requirement: required
  #     - :attr: abstract
  #       :xml_element: abstract
  #       :headings:
  #       - ABSTRACT
  #     - :attr: call_number
  #       :xml_element: callNumber
  #       :headings:
  #       - CALL NUMBER
  #     - :attr: citation_note
  #       :xml_element: citationNote
  #       :headings:
  #       - CITATION NOTE
  #     - :attr: collection_name
  #       :xml_element: collectionName
  #       :headings:
  #       - COLLECTION NAME
  #     - :attr: contributing_institution
  #       :xml_element: contributingInstitution
  #       :headings:
  #       - CONTRIBUTING INSTITUTION
  #     - :attr: contributor
  #       :xml_element: contributor
  #       :headings:
  #       - CONTRIBUTOR
  #       :multivalued: true
  #       :value_sep: '|'
  #     - :attr: corporate_name
  #       :xml_element: corporateName
  #       :headings:
  #       - CORPORATE NAME
  #       :multivalued: true
  #       :value_sep: '|'
  #     # ... etc.
  #
  # Spreadsheet <tt>ark+=99999=fk45t4vg3q/pqc_descriptive.xlsx</tt>:
  #
  #  | OBJECT TYPE | UNIQUE IDENTIFIER     | ABSTRACT                               | CALL NUMBER   | CITATION NOTE   | COLLECTION NAME                    | CONTRIBUTING INSTITUTION | CONTRIBUTOR           | CORPORATE NAME             | COVERAGE | CREATOR                                | DATE      | DESCRIPTION                   | EXTENT    | FORMAT | GEOGRAPHIC SUBJECT              | IDENTIFIER | INCLUDES          | LANGUAGE | MEDIUM | NOTES                   | PERSONAL NAME                                    | PROVENANCE   | PUBLISHER                 | RELATION           | RIGHTS                                             | SOURCE           | SUBJECT                                                         | TITLE                                                | DIRECTORY NAME | FILENAME(S)                     | STATUS |
  #  |-------------|-----------------------|----------------------------------------|---------------|-----------------|------------------------------------|--------------------------|-----------------------|----------------------------|----------|----------------------------------------|-----------|-------------------------------|-----------|--------|---------------------------------|------------|-------------------|----------|--------|-------------------------|--------------------------------------------------|--------------|---------------------------|--------------------|----------------------------------------------------|------------------|-----------------------------------------------------------------|------------------------------------------------------|----------------|---------------------------------|--------|
  #  | Manuscript  | ark:/99999/fk45t4vg3q | A collection of letters dealing [snip] | Ms. Coll. 637 | A citation note | Penn in Hand: Selected Manuscripts | U. Penn                  | Smith, J. | Brown, M. | Fleischman's Inc. | Macy's | Spain    | Anderson, A. | Berry, B. | Crowley, C. | 1459-1519 | The description of the object | 300 pages | Folio  | Barcelona, Spain | Milan, Italy | abcd-123   | Other Call number | Italian  | paper  | A note about the object | Grilinzone, Leone|Pio, Lionello|Burgo, Andrea de | Milan, Italy | Harcourt Brace Jovanovich | http://example.com | https://creativecommons.org/publicdomain/mark/1.0/ | manuscript image | Manuscripts, Renaissance|Manuscripts, Italian|Holy Roman Empire | Lettere latine di Alberto Pio Conte di Carpi, [snip] |                | file1.tif; file2.tif; file3.tif |        |
  #
  # You can generate the XML for the spreadsheet and directory as follows:
  #
  #     require 'yaml'
  #     require 'xlsx_to_pqc_xml'
  #
  #     sheet_config = YAML.load open('pqc_descriptive.xlsx').read
  #
  #     description_metadata = XlsxToPqcXml::DescriptiveMetadata.new package_directory: 'ark+=99999=fk45t4vg3q', sheet_config: sheet_config
  #     puts description_metadata.xml
  #
  # Will print the following XML:
  #
  #     <?xml version="1.0"?>
  #     <record>
  #       <ark>ark:/99999/fk45t4vg3q</ark>
  #       <pqc_elements>
  #         <pqc_element name="type">
  #           <value>Manuscript</value>
  #         </pqc_element>
  #         <pqc_element name="abstract">
  #           <value>A collection of letters dealing with foreign affairs of countries such as Lombardy, Spain, England, Flanders, and Naples, showing strong anti-French and anti-Venetian feeling. Half of the letters and memoranda are from Alberto Pio to Maximilian I or his officials; the other half are letters to Alberto Pio from, in order, Lionello Pio (his brother, 13 letters), Giovanni Matteo Giberti (bishop of Verona, 8 letters), Lorenzo Compeggi (1 letter), Andrea de Burgo (13 letters), Federico Fregoso (archbishop of Salerno, 1 letter), Leone Grilinzone (8 letters), Giovanni Battista Spinello (5 letters), and Jacopo Bannissi (20 letters). A few of the letters are written in code.</value>
  #         </pqc_element>
  #         <pqc_element name="callNumber">
  #           <value>Ms. Coll. 637</value>
  #         </pqc_element>
  #         <pqc_element name="citationNote">
  #           <value>A citation note</value>
  #         </pqc_element>
  #         <pqc_element name="collectionName">
  #           <value>Penn in Hand: Selected Manuscripts</value>
  #         </pqc_element>
  #         <pqc_element name="contributingInstitution">
  #           <value>U. Penn</value>
  #         </pqc_element>
  #         <pqc_element name="contributor">
  #           <value>Smith, J.</value>
  #           <value>Brown, M.</value>
  #         </pqc_element>
  #         <pqc_element name="corporateName">
  #           <value>Fleischman's Inc.</value>
  #           <value>Macy's</value>
  #         </pqc_element>
  #         <pqc_element name="coverage">
  #           <value>Spain</value>
  #         </pqc_element>
  #         <pqc_element name="creator">
  #           <value>Anderson, A.</value>
  #           <value>Berry, B.</value>
  #           <value>Crowley, C.</value>
  #         </pqc_element>
  #         <pqc_element name="date">
  #           <value>1459-1519</value>
  #         </pqc_element>
  #         <pqc_element name="description">
  #           <value>The description of the object</value>
  #         </pqc_element>
  #         <pqc_element name="extent">
  #           <value>300 pages</value>
  #         </pqc_element>
  #         <pqc_element name="format">
  #           <value>Folio</value>
  #         </pqc_element>
  #         <pqc_element name="geographicSubject">
  #           <value>Barcelona, Spain</value>
  #           <value>Milan, Italy</value>
  #         </pqc_element>
  #         <pqc_element name="identifier">
  #           <value>abcd-123</value>
  #         </pqc_element>
  #         <pqc_element name="includes">
  #           <value>Other Call number</value>
  #         </pqc_element>
  #         <pqc_element name="language">
  #           <value>Italian</value>
  #         </pqc_element>
  #         <pqc_element name="medium">
  #           <value>paper</value>
  #         </pqc_element>
  #         <pqc_element name="personalName">
  #           <value>Grilinzone, Leone</value>
  #           <value>Pio, Lionello</value>
  #           <value>Burgo, Andrea de</value>
  #         </pqc_element>
  #         <pqc_element name="provenance">
  #           <value>Milan, Italy</value>
  #         </pqc_element>
  #         <pqc_element name="publisher">
  #           <value>Harcourt Brace Jovanovich</value>
  #         </pqc_element>
  #         <pqc_element name="relation">
  #           <value>http://example.com</value>
  #         </pqc_element>
  #         <pqc_element name="rights">
  #           <value>https://creativecommons.org/publicdomain/mark/1.0/</value>
  #         </pqc_element>
  #         <pqc_element name="source">
  #           <value>manuscript image</value>
  #         </pqc_element>
  #         <pqc_element name="subject">
  #           <value>Manuscripts, Renaissance</value>
  #           <value>Manuscripts, Italian</value>
  #           <value>Holy Roman Empire</value>
  #         </pqc_element>
  #         <pqc_element name="title">
  #           <value>Lettere latine di Alberto Pio Conte di Carpi, Ambasciatore della Maesta&#x300; Cesarea in Roma.</value>
  #         </pqc_element>
  #       </pqc_elements>
  #     </record>
  #
  # === Spreadsheet configuration
  #
  # {XlsxToPqcXml::DescriptiveMetadata} uses a standard configuration as
  # described in the documentation for {XlsxToPqcXml::XlsxData} and adds
  # an +:xml_element+ key to the attribute configuration:
  #
  #     - :attr: contributing_institution
  #       :xml_element: contributingInstitution
  #       :headings:
  #       - CONTRIBUTING INSTITUTION
  #
  # Not all attributes need to have an +:xml_element+, but those that do
  # will be mapped to that value.
  #
  # Note that more than one attribute may have the same +:xml_element+. So, you
  # can do something like:
  #
  #      - :attr: description
  #        :xml_element: description
  #        :headings:
  #        - DESCRIPTION
  #      - :attr: address
  #        :xml_element: description
  #        :headings:
  #        - ADDRESS
  #
  # And have XML output like:
  #
  #         <pqc_element name="description">
  #           <value>The description of the object</value>
  #           <value>102 Main St., Springfield, USA</value>
  #         </pqc_element>
  #
  class DescriptiveMetadata

    # --- Default --- (in case we want to spreadsheet name make editable later)
    DEFAULT_PQC_DESCRIPTIVE_XLSX_BASE = 'pqc_descriptive.xlsx'.freeze

    ##
    # Create a new {DescriptiveMetadata} instance. The +package_directory+
    # should contain a spreadsheet named +pqc_descriptive.xlsx+. +sheet_config+
    # should a hash as below.
    #
    # {:sheet_name=>"Descriptive",
    #  :sheet_position=>0,
    #  :heading_type=>"row",
    #  :attributes=>
    #   [{:attr=>"object_type", :xml_element=>"type", :headings=>["OBJECT TYPE"], :requirement=>"required"},
    #    {:attr=>"unique_identifier", :xml_element=>"ark", :data_type=>:ark, :headings=>["UNIQUE IDENTIFIER"], :requirement=>"required"},
    #    {:attr=>"abstract", :xml_element=>"abstract", :headings=>["ABSTRACT"]},
    #    {:attr=>"call_number", :xml_element=>"callNumber", :headings=>["CALL NUMBER"]},
    #    {:attr=>"citation_note", :xml_element=>"citationNote", :headings=>["CITATION NOTE"]},
    #    {:attr=>"collection_name", :xml_element=>"collectionName", :headings=>["COLLECTION NAME"]},
    #    {:attr=>"contributing_institution", :xml_element=>"contributingInstitution", :headings=>["CONTRIBUTING INSTITUTION"]},
    #    {:attr=>"contributor", :xml_element=>"contributor", :headings=>["CONTRIBUTOR"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"corporate_name", :xml_element=>"corporateName", :headings=>["CORPORATE NAME"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"coverage", :xml_element=>"coverage", :headings=>["COVERAGE"]},
    #    {:attr=>"creator", :xml_element=>"creator", :headings=>["CREATOR"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"date", :xml_element=>"date", :headings=>["DATE"]},
    #    {:attr=>"description", :xml_element=>"description", :headings=>["DESCRIPTION"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"extent", :xml_element=>"extent", :headings=>["EXTENT"]},
    #    {:attr=>"format", :xml_element=>"format", :headings=>["FORMAT"]},
    #    {:attr=>"geographic_subject", :xml_element=>"geographicSubject", :headings=>["GEOGRAPHIC SUBJECT"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"identifier", :xml_element=>"identifier", :headings=>["IDENTIFIER"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"includes", :xml_element=>"includes", :headings=>["INCLUDES"]},
    #    {:attr=>"language", :xml_element=>"language", :headings=>["LANGUAGE"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"medium", :xml_element=>"medium", :headings=>["MEDIUM"]},
    #    {:attr=>"notes", :headings=>["NOTES"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"personal_name", :xml_element=>"personalName", :headings=>["PERSONAL NAME"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"provenance", :xml_element=>"provenance", :headings=>["PROVENANCE"]},
    #    {:attr=>"publisher", :xml_element=>"publisher", :headings=>["PUBLISHER"]},
    #    {:attr=>"relation", :xml_element=>"relation", :headings=>["RELATION"]},
    #    {:attr=>"rights", :xml_element=>"rights", :headings=>["RIGHTS"], :data_type=>:web_url},
    #    {:attr=>"source", :xml_element=>"source", :headings=>["SOURCE"]},
    #    {:attr=>"subject", :xml_element=>"subject", :headings=>["SUBJECT"], :multivalued=>true, :value_sep=>"|"},
    #    {:attr=>"title", :xml_element=>"title", :headings=>["TITLE"], :requirement=>"required"},
    #    {:attr=>"directory_name", :headings=>["DIRECTORY NAME"]},
    #    {:attr=>"filenames", :headings=>["FILENAME(S)"], :multivalued=>true, :value_sep=>";"},
    #    {:attr=>"status", :headings=>["STATUS"]}]}
    #
    def initialize package_directory: , sheet_config:
      @package_directory     = package_directory
      @sheet_config          = sheet_config
      @xlsx_data             = nil
      @descriptive_xlsx_base = DEFAULT_PQC_DESCRIPTIVE_XLSX_BASE
      @attribute_map         = {}
      @data_for_xml          = []
    end

    def xlsx_path
      File.join @package_directory, @descriptive_xlsx_base
    end

    def spreadsheet_data
      return @xlsx_data.data unless @xlsx_data.nil?

      @xlsx_data = XlsxData.new xlsx_path: xlsx_path, config: @sheet_config
      @xlsx_data.data
    end

    ##
    # Return a hash of attribute configurations, each mapped to its +:attr+.
    #
    # For example,
    #
    #   {
    #     :object_type=>{:attr=>"object_type", :xml_element=>"type", :headings=>["OBJECT TYPE"], :requirement=>"required"},
    #     :unique_identifier=>{:attr=>"unique_identifier", :xml_element=>"ark", :data_type=>:ark, :headings=>["UNIQUE IDENTIFIER"], :requirement=>"required"},
    #     # etc.
    #   }
    #
    #
    # @return [Hash]
    def attribute_map
      return @attribute_map unless @attribute_map.empty?

      @attribute_map = (@sheet_config[:attributes] || []).inject({}) do |memo,attr|
        memo[attr[:attr].to_sym] = attr
        memo
      end
    end

    ##
    # Extract the data from the spreadsheet, mapping each defined +:xml_element+
    # to an array of values. Returns an array of hashes, one hash for each
    # record in the spreadsheet.
    #
    # @return [Array<Hash>]
    def data_for_xml
      return @data_for_xml unless @data_for_xml.empty?

      spreadsheet_data.each do |record|
        @data_for_xml << record.inject({}) do |memo,attr_value|
          # get the XML element name for the attr
          element = attribute_map.dig attr_value.first, :xml_element
          next memo unless element
          value = attr_value.last
          next memo unless value
          ra = value.is_a?(Array) ?  value : [value]
          memo[element] ||= []
          memo[element] += ra
          memo
        end
      end

      @data_for_xml
    end

    def xml
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.records {
          data_for_xml.each do |record|
            xml.record {
              xml.ark record['ark'].first
              xml.pqc_elements {
                record.each do |key,values|
                  next if key == 'ark'
                  xml.pqc_element(name: key) {
                    values.each do |val|
                      xml.value val
                    end
                  }
                end
              }
            }
          end
        }
      end
      builder.to_xml
    end
  end
end
