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
    def size_and_position_image(image_path, x:, y:, max_height:, max_width:, v_align: :center, h_align: :center)
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

      target_y = case v_align
      when :center
        y + ((max_height - target_height) / 2.0).floor
      when :top
        y
      when :bottom
        y + max_height - target_height
      end

      target_x = case h_align
      when :center
        x + ((max_width - target_width) / 2.0).floor
      when :left
        x
      when :right
        x + max_width - target_width
      end

      Struct.new(:height, :width, :y, :x, keyword_init: true).new(
        height: target_height.round,
        width: target_width.round,
        y: target_y,
        x: target_x,
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
      result = remove_whitespace(result)
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

    # Remove whitespace at start of paragraphs while preserving attributes
    def remove_whitespace(ooxml)
      ooxml.gsub(/<a:t([^>]*)>[\s\t\n\r]*/,'<a:t\1>')
    end

    def remove_declaration(ooxml)
      ooxml.sub(/<\?xml (.*?)>/, '').gsub(/\s*xmlns:(\w+)="(.*?)\s*"/, '')
    end

    def remove_newlines(ooxml)
      ooxml.gsub("\n","")
    end

    def extract_ooxml_text(ooxml)
      if ooxml.is_a?(String)
        ooxml = Nokogiri::XML.fragment(ooxml)
      end

      if ooxml.children.any?
        child_text = ooxml.children.map { |child|
          str = extract_ooxml_text(child).gsub(/[[:space:]]/, ' ').strip
          str || ''
        }

        # If paragraph, inline text
        # Otherwise, join with newlines
        if ooxml.name == 'a:p'
          child_text.reject { |str| str.empty? }.join(' ')
        else
          child_text.join("\n")
        end
      else
        ooxml.text
      end
    end

    def measure_text(
      text,
      font_family: 'Arial',
      font_size: 14,
      font_style: Magick::NormalStyle,
      font_weight: Magick::NormalWeight
    )
      unless text && text.length > 0
        return { width: 0, height: 0 }
      end

      label = Magick::Draw.new
      label.font = font_family
      label.pointsize = font_size
      label.text_antialias(true)
      label.font_style = font_style
      label.font_weight = font_weight
      label.gravity = Magick::CenterGravity

      text.split(/\n/).reduce({ width: 0, height: 0 }) do |d, line|
        label.text(0, 0, text)
        metrics = label.get_type_metrics(text)
        width = metrics.width
        height = metrics.height

        {
          width: [d[:width], width].max,
          height: d[:height] + height
        }
      end
    end
  end
end
