RSpec.configure do |rspec|
  # This config option will be enabled by default on RSpec 4,
  # but for reasons of backwards compatibility, you have to
  # set it on RSpec 3.
  #
  # It causes the host group and examples to inherit metadata
  # from the shared context.
  rspec.shared_context_metadata_behavior = :apply_to_host_groups
end

RSpec.shared_context 'shared context', :shared_context => :metadata do
  let(:fixtures_path) { File.expand_path '../../fixtures', __FILE__ }
  let(:tmp_dir) { File.expand_path '../../../tmp', __FILE__ }
  let(:staging_dir) { File.join tmp_dir, 'staging'}

end

RSpec.configure do |rspec|
  rspec.include_context 'shared context', :include_shared => true
end
