require 'spec_helper'

include XlsxToPqcXml

RSpec.describe StructuralMetadata do
  include_context 'shared context'

  let(:tiff_files) do
    %w{ 0001.tif
        0002.tif
        0002a.tif
        0002b.tif
        0003.tif
        0004.tif
        0005.tif
        0006.tif
        0007.tif
        reference.tif }.freeze
  end

  let(:ark)                   { 'ark:/99999/fk42244n9f' }
  let(:ark_dir)               { ark.gsub(%r{:}, '+').gsub(%r{/}, '=') }

  let(:good_structural_xlsx)  { File.join fixtures_dir, 'good_pqc_structural.xlsx' }
  let(:package_dir)           { File.join staging_dir, ark_dir }
  let(:structural_config_yml) { File.join fixtures_dir, 'structural_config.yml' }
  let(:sheet_config)          { YAML.load open(structural_config_yml).read }

  let(:structural_xlsx)       { File.join package_dir, 'pqc_structural.xlsx' }
  let(:structural_metadata)   { StructuralMetadata.new package_directory: package_dir, sheet_config: sheet_config }

  # let(:multiple_structural)

  after :each do
    FileUtils.remove_dir staging_dir if File.exists? staging_dir
  end

  context 'new' do
    it 'should create a StructuralMetadata instance' do
      expect(StructuralMetadata.new package_directory: 'path', sheet_config: sheet_config).to be_a StructuralMetadata
    end
  end

  context '#spreadsheet_data' do
    before :each do
      FileUtils.remove_dir staging_dir if File.exists? staging_dir
      FileUtils.mkdir_p package_dir    unless File.exists? package_dir
      FileUtils.cp good_structural_xlsx, structural_xlsx
    end

    let(:spreadsheet_data) { structural_metadata.spreadsheet_data }

    it 'should return an array of hashes' do
      expect(spreadsheet_data).to be_an Array
      expect(spreadsheet_data.first).to be_a Hash
    end
  end

  context '#data_file_list' do
    before :each do
      FileUtils.remove_dir staging_dir if File.exists? staging_dir
      FileUtils.mkdir_p package_dir    unless File.exists? package_dir
      FileUtils.cp good_structural_xlsx, structural_xlsx
      FileUtils.touch tiff_files.map { |t| File.join package_dir, t }
    end

    let(:files_on_disk) { structural_metadata.files_on_disk }

    it 'should list all the tiffs' do
      expect(files_on_disk.sort).to eq tiff_files.sort
    end
  end

  context '#image_data' do
    before :each do
      FileUtils.remove_dir staging_dir if File.exists? staging_dir
      FileUtils.mkdir_p package_dir    unless File.exists? package_dir
      FileUtils.cp good_structural_xlsx, structural_xlsx
      FileUtils.touch tiff_files.map { |t| File.join package_dir, t }
    end

    let(:image_data) { structural_metadata.image_data }

    it 'should return an array of OpenStructs' do
      expect(image_data).to be_an Array
      expect(image_data.first).to be_an OpenStruct
    end
  end

  context '#xml' do
    before :each do
      FileUtils.remove_dir staging_dir if File.exists? staging_dir
      FileUtils.mkdir_p package_dir    unless File.exists? package_dir
      FileUtils.cp good_structural_xlsx, structural_xlsx
      FileUtils.touch tiff_files.map { |t| File.join package_dir, t }
    end

    let(:parsed_xml) { Nokogiri::XML structural_metadata.xml }

    it 'should should return a string' do
      expect(structural_metadata.xml).to be_a String
    end

    it 'should be valid xml' do
      expect(parsed_xml).to be_a Nokogiri::XML::Document
    end

    it 'should have one page for each image' do
      expect(parsed_xml).to have_xpath('//pages/page', count: 10)
    end

    it 'should have two toc entries' do
      expect(parsed_xml).to have_xpath('//tocentry[@name="toc"]', count: 2)
    end

    it 'should have two ill entries' do
      expect(parsed_xml).to have_xpath('//tocentry[@name="ill"]', count: 2)
    end

    # <page number="6" seq="6" image.defaultscale="3" side="verso" id="0006" image.id="0006" visiblepage="3v-4r" display="true">
    it 'should have a page @number' do
      expect(parsed_xml).to have_xpath('//page[@number=6]', count: 1)
    end

    it 'should have a page @seq' do
      expect(parsed_xml).to have_xpath('//page[@seq=6]', count: 1)
    end

    it 'should have an @image.defaultscale' do
      # there are 10 images in :tiff_files
      expect(parsed_xml).to have_xpath('//page[@image.defaultscale=3]', count: 10)
    end

    it 'should have a page @side' do
      expect(parsed_xml).to have_xpath('//page[@side="verso"]', count: 5)
    end

    it 'should have an image @id' do
      expect(parsed_xml).to have_xpath('//page[@id="0001"]', count: 1)
    end

    it 'should have an @image.id' do
      expect(parsed_xml).to have_xpath('//page[@image.id="0001"]', count: 1)
    end

    it 'should have a @visiblepage' do
      expect(parsed_xml).to have_xpath('//page[@visiblepage="3v-4r"]', count: 1)
    end

    it 'should have @display = true the images in the spreadsheet' do
      # 7 of the 10 images are in the spreadsheet
      expect(parsed_xml).to have_xpath('//page[@display="true"]', count: 7)
    end
  end

end
