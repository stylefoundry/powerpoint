require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'
require 'sanitize'

module Powerpoint
  module Slide
    class FFTrendWhatNext
      include Powerpoint::Util

      attr_reader :title, :content, :cols

      def initialize(options={})
        require_arguments [:presentation, :title, :content], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}

        @cols = []
        for i in 0..2
           @cols[i] = []
            if content["rowsManagerInput"]["value"][0] != nil && content["rowsManagerInput"]["value"][0]["item"]["items"]["textInput#{i+1}"] != nil
              @cols[i] << Sanitize.clean(content["rowsManagerInput"]["value"][0]["item"]["items"]["textInput#{i+1}"]["value"])
           end
            if content["rowsManagerInput"]["value"][1] != nil && content["rowsManagerInput"]["value"][1]["item"]["items"]["textInput#{i+1}"] != nil
              @cols[i] << Sanitize.clean(content["rowsManagerInput"]["value"][1]["item"]["items"]["textInput#{i+1}"]["value"])
            end
            if content["rowsManagerInput"]["value"][2] != nil && content["rowsManagerInput"]["value"][2]["item"]["items"]["textInput#{i+1}"] != nil
              @cols[i] << Sanitize.clean(content["rowsManagerInput"]["value"][2]["item"]["items"]["textInput#{i+1}"]["value"])
            end
        end

      end

      def save(extract_path, index)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        nil
      end

      def save_rel_xml(extract_path, index)
        render_view('ff_trend_what_next_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_trend_what_next_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end
