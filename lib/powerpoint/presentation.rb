require 'zip/filesystem'
require 'fileutils'
require 'tmpdir'

module Powerpoint
  class Presentation
    include Powerpoint::Util

    attr_reader :slides

    def initialize
      @slides = []
    end

    def add_intro(title, subtitile = nil)
      existing_intro_slide = @slides.select {|s| s.class == Powerpoint::Slide::Intro}[0]
      slide = Powerpoint::Slide::Intro.new(presentation: self, title: title, subtitile: subtitile)
      if existing_intro_slide
        @slides[@slides.index(existing_intro_slide)] = slide
      else
        @slides.insert 0, slide
      end
    end

    def add_textual_slide(title, content = [])
      @slides << Powerpoint::Slide::Textual.new(presentation: self, title: title, content: content)
    end

    def add_pictorial_slide(title, image_path, coords = {})
      @slides << Powerpoint::Slide::Pictorial.new(presentation: self, title: title, image_path: image_path, coords: coords)
    end

    def add_text_picture_slide(title, image_path, content = [])
      @slides << Powerpoint::Slide::TextPicSplit.new(presentation: self, title: title, image_path: image_path, content: content)
    end

    def add_picture_description_slide(title, image_path, content = [])
      @slides << Powerpoint::Slide::DescriptionPic.new(presentation: self, title: title, image_path: image_path, content: content)
    end

    def add_ff_trend_intro_slide(title, subtitle, image_path, coords = {})
      existing_intro_slide = @slides.select {|s| s.class == Powerpoint::Slide::Intro}[0]
      slide = Powerpoint::Slide::FFTrendIntro.new(presentation: self, title: title, subtitle: subtitle, image_path: image_path, coords: {})
      if existing_intro_slide
        @slides[@slides.index(existing_intro_slide)] = slide
      else
        @slides.insert 0, slide
      end
    end

    def add_ff_what_next_slide(title, content = {})
      @slides << Powerpoint::Slide::FFTrendWhatNext.new(presentation: self, title: title, content: content)
    end

    def add_ff_heading_text_slide(title, content)
      @slides << Powerpoint::Slide::FFTrendHeadingText.new(presentation: self, title: title, content: content)
    end

    def add_ff_three_row_text_slide(title, content)
      @slides << Powerpoint::Slide::FFTrendThreeRowText.new(presentation: self, title: title, content: content)
    end

    def add_ff_sector_impact_slide(title, content)
      @slides << Powerpoint::Slide::FFTrendSectorImpact.new(presentation: self, title: title, content: content)
    end

    def add_ff_associated_content_slide(title, subtitle, image_path, coords = {}, link_path = nil)
      @slides << Powerpoint::Slide::FFTrendIntro.new(presentation: self, title: title, subtitle: subtitle, image_path: image_path,  coords: {}, link_path: link_path)
    end

    def add_ff_embeded_slide(slide_content, slide_rel_content, images, charts, embeddings)
      @slides << Powerpoint::Slide::FFEmbededSlide.new(presentation: self, title: "", content: slide_content, rel_content: slide_rel_content, images: images, charts: charts, embeddings: embeddings)
    end

    def save(path)
      Dir.mktmpdir do |dir|
        extract_path = "#{dir}/extract_#{Time.now.strftime("%Y-%m-%d-%H%M%S")}"

        # Copy template to temp path
        FileUtils.copy_entry(TEMPLATE_PATH, extract_path)

        # Remove keep files
        Dir.glob("#{extract_path}/**/.keep").each do |keep_file|
          FileUtils.rm_rf(keep_file)
        end

        # Render/save generic stuff
        render_view('content_type.xml.erb', "#{extract_path}/[Content_Types].xml")
        render_view('presentation.xml.rel.erb', "#{extract_path}/ppt/_rels/presentation.xml.rels")
        render_view('presentation.xml.erb', "#{extract_path}/ppt/presentation.xml")
        #render_view('view_props.xml.erb', "#{extract_path}/ppt/viewProps.xml")
        #render_view('table_styles.xml.erb', "#{extract_path}/ppt/tableStyles.xml")
        #render_view('pres_props.xml.erb', "#{extract_path}/ppt/presProps.xml")
        render_view('app.xml.erb', "#{extract_path}/docProps/app.xml")
        #render_view('core.xml.erb', "#{extract_path}/docProps/core.xml")

        # Save slides
        slides.each_with_index do |slide, index|
          slide.save(extract_path, index + 1)
        end

        # Create .pptx file
        File.delete(path) if File.exist?(path)
        Powerpoint.compress_pptx(extract_path, path)
      end

      path
    end

    def file_types
      slides.map {|slide| slide.file_type if slide.respond_to? :file_type }.compact.uniq
    end
  end
end
