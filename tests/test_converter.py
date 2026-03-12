"""Tests for tomd.converter."""

import pytest
from tomd.converter import strip_base64_images


class TestStripBase64Images:
    def test_replaces_base64_png(self):
        text = '![alt text](data:image/png;base64,iVBORw0KGgoAAAANSUhEUg==)'
        result = strip_base64_images(text)
        assert result == '![alt text]()'

    def test_replaces_base64_jpeg(self):
        text = '![photo](data:image/jpeg;base64,/9j/4AAQSkZJRg==)'
        result = strip_base64_images(text)
        assert result == '![photo]()'

    def test_preserves_empty_alt(self):
        text = '![](data:image/png;base64,abc123)'
        result = strip_base64_images(text)
        assert result == '![]()'

    def test_preserves_normal_images(self):
        text = '![alt](https://example.com/image.png)'
        result = strip_base64_images(text)
        assert result == '![alt](https://example.com/image.png)'

    def test_preserves_plain_text(self):
        text = 'Hello world'
        result = strip_base64_images(text)
        assert result == 'Hello world'

    def test_multiple_images(self):
        text = (
            '![a](data:image/png;base64,AAA) middle '
            '![b](data:image/gif;base64,BBB)'
        )
        result = strip_base64_images(text)
        assert result == '![a]() middle ![b]()'
