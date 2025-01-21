require 'htmltoooxml'
require 'rmagick'
require 'pry'

module Powerpoint
  module Util
    PT_CONVERSION_FACTOR = 12700

    def pixle_to_pt(px)
      px * PT_CONVERSION_FACTOR
    end

    def pt_to_pixle(pt)
      pt / PT_CONVERSION_FACTOR
    end

    # Get largest dimensions that maintain image aspect ratio and fit inside max
    # Returns x and y such that element is centered in container
    def size_and_position_image(image_path, x:, y:, max_height:, max_width:)
      image = Magick::ImageList.new(image_path).first
      image_height = image.rows
      image_width = image.columns
      image_ratio = image_width / image_height.to_f
      # Maximise image height
      target_height = max_height
      target_width = target_height * image_ratio

      # Image height is less constrained than width, swap to maximise width
      if target_width > max_width
        target_width = max_width
        target_height = target_width / image_ratio
      end

      Struct.new(:height, :width, :y, :x, keyword_init: true).new(
        height: target_height.round,
        width: target_width.round,
        # Half of remaining height (Round down)
        y: y + ((max_height - target_height) / 2.0).floor,
        # Half of remaining width (Round down)
        x: x + ((max_width - target_width) / 2.0).floor,
      )
    end

    def render_view(template_name, path)
      view_contents = read_template(template_name)
      renderer = ERB.new(view_contents)
      data = renderer.result(binding)
      File.open(path, 'wb:UTF-8') { |f| f << data }
    end

    def render_raw(data, path)
      File.open(path, 'wb:UTF-8') { |f| f << data }
    end

    def read_template(filename)
      File.read("#{Powerpoint::VIEW_PATH}/#{filename}")
    end

    def require_arguments(required_arguments, arguments)
      missing = required_arguments - arguments.keys
      raise ArgumentError, "Missing required arguments: #{missing.join(', ')}" unless missing.empty?
    end

    def copy_media(extract_path, image_path)
      image_name = File.basename(image_path)
      dest_path = "#{extract_path}/ppt/media/#{image_name}"
      FileUtils.copy_file(image_path, dest_path) unless File.exist?(dest_path)
    end

    def html_to_ooxml(html)
      source = Nokogiri::HTML(html.gsub(/>\s+</, '><'))
      result = Htmltoooxml::Document.new().transform_doc_xml(source, false)
      result.gsub!(/\s*<!--(.*?)-->\s*/m, '')
      result = remove_declaration(result)
      result = remove_newlines(result)
    end

    def ooxml_blank?(ooxml)
      if ooxml.nil?
        true
      else
        is_blank_node = -> (node) do
          case node
          when Nokogiri::XML::Text
            /\A[[:space:]]*\z/ =~ node.content
          else
            node.children.all?(&is_blank_node)
          end
        end

        Nokogiri::XML.parse(ooxml).children.all?(&is_blank_node)
      end
    end

    def remove_whitespace(ooxml)
      ooxml.gsub(/\s+/, ' ').gsub(/>\s+</, '><').strip
    end

    def remove_declaration(ooxml)
      ooxml.sub(/<\?xml (.*?)>/, '').gsub(/\s*xmlns:(\w+)="(.*?)\s*"/, '')
    end

    def remove_newlines(ooxml)
      ooxml.gsub("\n","")
    end
  end
end
