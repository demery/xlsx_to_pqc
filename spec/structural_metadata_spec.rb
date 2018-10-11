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

  let(:good_structural_xlsx)  { File.join fixtures_path, 'good_pqc_structural.xlsx' }
  let(:package_dir)           { File.join staging_dir, ark_dir }

  let(:structural_xlsx)       { File.join package_dir, 'pqc_structural.xlsx' }
  let(:structural_metadata)   { StructuralMetadata.new package_dir }

  after :each do
    FileUtils.remove_dir staging_dir if File.exists? staging_dir
  end

  context 'new' do
    it 'should create a StructuralMetadata instance' do
      expect(StructuralMetadata.new 'path').to be_a StructuralMetadata
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
      # spreadsheet_data = structural_metadata.spreadsheet_data
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

end
