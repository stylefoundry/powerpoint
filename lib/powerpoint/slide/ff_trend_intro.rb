require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class FFTrendIntro
      include Powerpoint::Util

      attr_reader :image_name, :title, :subtitle, :coords, :image_path, :link_path

      def initialize(options={})
        require_arguments [:presentation, :title, :subtitle, :image_path], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @coords = default_coords unless @coords.any?
        @image_name = File.basename(@image_path) if @image_path != nil
      end

      def save(extract_path, index)
        copy_media(extract_path, @image_path) if @image_path != nil
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        File.extname(image_name).gsub('.', '')
      end

      def default_coords
        slide_width = pixle_to_pt(720)
        default_width = pixle_to_pt(720)
        default_height = pixle_to_pt(481)

        return {} unless dimensions = FastImage.size(image_path)
        image_width, image_height = dimensions.map {|d| pixle_to_pt(d)}
        new_width = default_width < image_width ? default_width : image_width
        ratio = new_width / image_width.to_f
        new_height = default_height #(image_height.to_f * ratio).round
        {x: 0, y: 0, cx: new_width, cy: new_height}
      end
      private :default_coords

      def save_rel_xml(extract_path, index)
        render_view('ff_trend_intro_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_trend_intro_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end
