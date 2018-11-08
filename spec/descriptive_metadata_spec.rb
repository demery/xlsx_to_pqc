require 'spec_helper'

include XlsxToPqcXml

RSpec.describe StructuralMetadata do
  include_context 'shared context'


  let(:ark)                   { 'ark:/99999/fk42244n9f' }
  let(:ark_dir)               { ark.gsub(%r{:}, '+').gsub(%r{/}, '=') }

  let(:valid_descriptive_xlsx)  { fixture_path 'valid_pqc_descriptive.xlsx' }
  let(:package_dir)           { File.join staging_dir, ark_dir }
  let(:descriptive_config_yml) { fixture_path 'descriptive_config.yml' }
  let(:sheet_config)          { YAML.load open(descriptive_config_yml).read }

  let(:descriptive_xlsx)       { File.join package_dir, 'pqc_descriptive.xlsx' }
  let(:descriptive_metadata)   { DescriptiveMetadata.new package_directory: package_dir, sheet_config: sheet_config }

  after :each do
    FileUtils.remove_dir staging_dir if File.exists? staging_dir
  end

  context 'new' do
    it 'should create a DescriptiveMetadata instance' do
      expect(DescriptiveMetadata.new package_directory: 'path', sheet_config: sheet_config).to be_a DescriptiveMetadata
    end
  end

  context '#spreadsheet_data' do
    before :each do
      FileUtils.remove_dir staging_dir if File.exists? staging_dir
      FileUtils.mkdir_p package_dir    unless File.exists? package_dir
      FileUtils.cp valid_descriptive_xlsx, descriptive_xlsx
    end

    let(:spreadsheet_data) { descriptive_metadata.spreadsheet_data }

    it 'should return an array of hashes' do
      expect(spreadsheet_data).to be_an Array
      expect(spreadsheet_data.first).to be_a Hash
    end
  end

  context '#attribute_map' do
    before :each do
      FileUtils.remove_dir staging_dir if File.exists? staging_dir
      FileUtils.mkdir_p package_dir    unless File.exists? package_dir
      FileUtils.cp valid_descriptive_xlsx, descriptive_xlsx
    end

    it 'should return a hash' do
      expect(descriptive_metadata.attribute_map).to be_a Hash
    end
  end

  context '#data_for_xml' do
    before :each do
      FileUtils.remove_dir staging_dir if File.exists? staging_dir
      FileUtils.mkdir_p package_dir    unless File.exists? package_dir
      FileUtils.cp valid_descriptive_xlsx, descriptive_xlsx
    end

    it 'should return an array of hashes' do
      expect(descriptive_metadata.data_for_xml).to be_a Array
      expect(descriptive_metadata.data_for_xml.first).to be_a Hash
    end
  end

  context '#xml' do
    before :each do
      FileUtils.remove_dir staging_dir if File.exists? staging_dir
      FileUtils.mkdir_p package_dir    unless File.exists? package_dir
      FileUtils.cp valid_descriptive_xlsx, descriptive_xlsx
    end

    let(:parsed_xml) { Nokogiri::XML descriptive_metadata.xml }

    it 'should should return a string' do
      expect(descriptive_metadata.xml).to be_a String
    end

    it 'should be valid xml' do
      expect(parsed_xml).to be_a Nokogiri::XML::Document
    end

    it 'should have one pqc_elements tag' do
      expect(parsed_xml).to have_xpath('//pqc_elements', count: 1)
    end

    it 'should have pqc_element tags' do
      expect(parsed_xml).to have_xpath('//pqc_elements/pqc_element', count: 27)
    end

    it 'should have a type' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="type"]/value', count: 1)
    end

    it 'should have an abstract' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="abstract"]/value', count: 1)
    end

    it 'should have a callNumber' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="callNumber"]/value', count: 1)
    end

    it 'should have a citationNote' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="citationNote"]/value', count: 1)
    end

    it 'should have a collectionName' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="collectionName"]/value', count: 1)
    end

    it 'should have a contributingInstitution' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="contributingInstitution"]/value', count: 1)
    end

    it 'should have 2 contributors' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="contributor"]/value', count: 2)
    end

    it 'should have 2 corporateNames' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="corporateName"]/value', count: 2)
    end

    it 'should have 2 corporateNames' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="corporateName"]/value', count: 2)
    end

    it 'should have a coverage' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="coverage"]/value', count: 1)
    end

    it 'should have 3 creators' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="creator"]/value', count: 3)
    end

    it 'should have a date' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="date"]')
    end

    it 'should have a description' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="description"]', count: 1)
    end

    it 'should have a extent' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="extent"]', count: 1)
    end

    it 'should have a format' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="format"]', count: 1)
    end

    it 'should have a geographicSubject' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="geographicSubject"]', count: 1)
    end

    it 'should have a identifier' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="identifier"]', count: 1)
    end

    it 'should have a includes' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="includes"]', count: 1)
    end

    it 'should have a language' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="language"]', count: 1)
    end

    it 'should have a medium' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="medium"]', count: 1)
    end

    it 'should have a personalName' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="personalName"]', count: 1)
    end

    it 'should have a provenance' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="provenance"]', count: 1)
    end

    it 'should have a publisher' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="publisher"]', count: 1)
    end

    it 'should have a relation' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="relation"]', count: 1)
    end

    it 'should have a rights' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="rights"]', count: 1)
    end

    it 'should have a source' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="source"]', count: 1)
    end

    it 'should have a subject' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="subject"]', count: 1)
    end

    it 'should have a title' do
      expect(parsed_xml).to have_xpath('//pqc_element[@name="title"]', count: 1)
    end
  end

end
