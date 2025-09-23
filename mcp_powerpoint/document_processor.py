#!/usr/bin/env python3
"""
Document processing for RAG system
Handles multiple file formats and extracts semantic content
"""

import asyncio
import aiofiles
import os
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path
import PyPDF2
import docx
from sentence_transformers import SentenceTransformer
import nltk
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk.corpus import stopwords
from nltk.chunk import ne_chunk
from nltk.tag import pos_tag
from collections import Counter
import json

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DocumentProcessor:
    """Advanced document processing for RAG system"""

    def __init__(self, model_name: str = 'all-MiniLM-L6-v2'):
        self.model_name = model_name
        self.sentence_model = None
        self._ensure_nltk_data()
        self.stopwords = set(stopwords.words('english'))

    def _ensure_nltk_data(self):
        """Ensure required NLTK data is downloaded"""
        required_data = ['punkt', 'averaged_perceptron_tagger',
                        'maxent_ne_chunker', 'words', 'stopwords']
        for data in required_data:
            try:
                nltk.data.find(f'tokenizers/{data}')
            except LookupError:
                try:
                    nltk.download(data, quiet=True)
                except Exception as e:
                    logger.warning(f"Failed to download NLTK data {data}: {e}")

    async def _get_sentence_model(self):
        """Lazy load sentence transformer model"""
        if self.sentence_model is None:
            self.sentence_model = SentenceTransformer(self.model_name)
        return self.sentence_model

    async def process_document(self, file_path: str) -> List[Dict[str, Any]]:
        """
        Process document into chunks with metadata

        Args:
            file_path: Path to document file

        Returns:
            List of processed chunks with metadata
        """
        try:
            logger.info(f"Processing document: {file_path}")

            # Extract text based on file type
            text_content = await self._extract_text(file_path)

            if not text_content.strip():
                logger.warning(f"No text content extracted from {file_path}")
                return []

            # Split into sentences
            sentences = sent_tokenize(text_content)

            # Create chunks (group sentences into chunks of ~200 words)
            chunks = self._create_semantic_chunks(sentences)

            # Process each chunk
            processed_chunks = []
            for i, chunk in enumerate(chunks):
                try:
                    chunk_data = {
                        "text": chunk,
                        "chunk_id": i,
                        "source_file": file_path,
                        "word_count": len(word_tokenize(chunk)),
                        "entities": self._extract_entities(chunk),
                        "key_terms": self._extract_key_terms(chunk),
                        "summary": await self._summarize_chunk(chunk)
                    }
                    processed_chunks.append(chunk_data)
                except Exception as e:
                    logger.error(f"Error processing chunk {i} from {file_path}: {e}")
                    continue

            logger.info(f"Successfully processed {len(processed_chunks)} chunks from {file_path}")
            return processed_chunks

        except Exception as e:
            logger.error(f"Error processing document {file_path}: {e}")
            return []

    async def _extract_text(self, file_path: str) -> str:
        """Extract text from various file formats"""
        file_ext = os.path.splitext(file_path)[1].lower()

        try:
            if file_ext == '.pdf':
                return await self._extract_pdf_text(file_path)
            elif file_ext == '.docx':
                return await self._extract_docx_text(file_path)
            elif file_ext in ['.txt', '.md']:
                return await self._extract_plain_text(file_path)
            else:
                logger.warning(f"Unsupported file type: {file_ext}")
                return ""
        except Exception as e:
            logger.error(f"Error extracting text from {file_path}: {e}")
            return ""

    async def _extract_pdf_text(self, file_path: str) -> str:
        """Extract text from PDF file"""
        text = ""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
        except Exception as e:
            logger.error(f"Error reading PDF {file_path}: {e}")
        return text

    async def _extract_docx_text(self, file_path: str) -> str:
        """Extract text from DOCX file"""
        try:
            doc = docx.Document(file_path)
            return "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()])
        except Exception as e:
            logger.error(f"Error reading DOCX {file_path}: {e}")
            return ""

    async def _extract_plain_text(self, file_path: str) -> str:
        """Extract text from plain text file"""
        try:
            async with aiofiles.open(file_path, 'r', encoding='utf-8') as file:
                return await file.read()
        except UnicodeDecodeError:
            # Try with different encoding
            try:
                async with aiofiles.open(file_path, 'r', encoding='latin-1') as file:
                    return await file.read()
            except Exception as e:
                logger.error(f"Error reading text file {file_path}: {e}")
                return ""
        except Exception as e:
            logger.error(f"Error reading text file {file_path}: {e}")
            return ""

    def _create_semantic_chunks(self, sentences: List[str],
                              target_chunk_size: int = 200) -> List[str]:
        """Create semantic chunks from sentences"""
        chunks = []
        current_chunk = []
        current_word_count = 0

        for sentence in sentences:
            if not sentence.strip():
                continue

            sentence_word_count = len(word_tokenize(sentence))

            if current_word_count + sentence_word_count > target_chunk_size and current_chunk:
                chunks.append(" ".join(current_chunk))
                current_chunk = [sentence]
                current_word_count = sentence_word_count
            else:
                current_chunk.append(sentence)
                current_word_count += sentence_word_count

        if current_chunk:
            chunks.append(" ".join(current_chunk))

        return chunks

    def _extract_entities(self, text: str) -> List[str]:
        """Extract named entities from text"""
        try:
            tokens = word_tokenize(text)
            pos_tags = pos_tag(tokens)
            tree = ne_chunk(pos_tags)

            entities = []
            for subtree in tree:
                if hasattr(subtree, 'label'):
                    entity = " ".join([token for token, pos in subtree.leaves()])
                    entities.append(entity)

            return entities
        except Exception as e:
            logger.error(f"Error extracting entities: {e}")
            return []

    def _extract_key_terms(self, text: str) -> List[str]:
        """Extract key terms from text"""
        try:
            tokens = word_tokenize(text.lower())

            # Filter out stopwords and punctuation
            filtered_tokens = [
                token for token in tokens
                if token.isalnum() and token not in self.stopwords and len(token) > 2
            ]

            # Get POS tags and keep only nouns, verbs, and adjectives
            pos_tags = pos_tag(filtered_tokens)
            key_terms = [
                token for token, pos in pos_tags
                if pos.startswith(('NN', 'VB', 'JJ'))
            ]

            # Return top terms by frequency
            term_counts = Counter(key_terms)
            return [term for term, count in term_counts.most_common(10)]
        except Exception as e:
            logger.error(f"Error extracting key terms: {e}")
            return []

    async def _summarize_chunk(self, chunk: str, max_length: int = 100) -> str:
        """Generate summary of chunk"""
        try:
            # Simple summarization: return first sentence or truncated text
            sentences = sent_tokenize(chunk)
            if sentences:
                first_sentence = sentences[0]
                if len(first_sentence) <= max_length:
                    return first_sentence
                else:
                    return first_sentence[:max_length-3] + "..."
            else:
                return chunk[:max_length-3] + "..." if len(chunk) > max_length else chunk
        except Exception as e:
            logger.error(f"Error summarizing chunk: {e}")
            return chunk[:max_length-3] + "..." if len(chunk) > max_length else chunk

    async def process_directory(self, directory_path: str,
                              file_patterns: List[str] = None) -> Dict[str, List[Dict[str, Any]]]:
        """
        Process all documents in a directory

        Args:
            directory_path: Directory containing documents
            file_patterns: File patterns to match

        Returns:
            Dictionary mapping file paths to processed chunks
        """
        if not file_patterns:
            file_patterns = ["*.txt", "*.md", "*.pdf", "*.docx"]

        directory_path = Path(directory_path)
        if not directory_path.exists():
            logger.error(f"Directory does not exist: {directory_path}")
            return {}

        results = {}

        for pattern in file_patterns:
            files = list(directory_path.glob(pattern))

            for file_path in files:
                try:
                    chunks = await self.process_document(str(file_path))
                    if chunks:
                        results[str(file_path)] = chunks
                except Exception as e:
                    logger.error(f"Error processing file {file_path}: {e}")
                    continue

        logger.info(f"Processed {len(results)} files from {directory_path}")
        return results

    async def generate_embeddings(self, texts: List[str]) -> List[List[float]]:
        """Generate embeddings for texts using sentence transformer"""
        try:
            model = await self._get_sentence_model()
            embeddings = model.encode(texts, convert_to_tensor=False)
            return embeddings.tolist() if hasattr(embeddings, 'tolist') else embeddings
        except Exception as e:
            logger.error(f"Error generating embeddings: {e}")
            return []


async def main():
    """Test the document processor"""
    processor = DocumentProcessor()

    # Test with a simple text file
    test_content = """This is a test document for the PowerPoint MCP tools.
    The document contains information about business presentations and data analysis.
    Companies often use PowerPoint to communicate strategic goals and project updates.
    """

    # Create a temporary test file
    test_file = "test_document.txt"
    async with aiofiles.open(test_file, 'w') as f:
        await f.write(test_content)

    try:
        # Process the test document
        chunks = await processor.process_document(test_file)
        print(f"Processed {len(chunks)} chunks:")
        for i, chunk in enumerate(chunks):
            print(f"\nChunk {i}:")
            print(f"  Text: {chunk['text'][:100]}...")
            print(f"  Entities: {chunk['entities']}")
            print(f"  Key terms: {chunk['key_terms']}")
            print(f"  Summary: {chunk['summary']}")
    finally:
        # Clean up test file
        if os.path.exists(test_file):
            os.unlink(test_file)


if __name__ == "__main__":
    asyncio.run(main())