require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class FFTrendSectorImpact
      include Powerpoint::Util

      attr_reader :title, :content, :image_path, :image_name, :links

      def initialize(options={})
        require_arguments [:presentation, :title, :content, :image_path, :links], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @image_name = File.basename(@image_path) if @image_path != nil
      end

      def save(extract_path, index)
        copy_media(extract_path, @image_path) if @image_path != nil
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        [{ type: MimeMagic.by_magic(File.open(image_path)).type, path: "/ppt/media/#{image_name}" }]
      end

      def save_rel_xml(extract_path, index)
        render_view('ff_trend_sector_impact_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_trend_sector_impact_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end
