package com.bj.tika.extraction;

import java.io.IOException;
import java.io.InputStream;

import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.sax.BodyContentHandler;
import org.xml.sax.SAXException;


public class AutoDetectionExample {
    public static void main(final String[] args) throws IOException,
            SAXException, TikaException {
        Parser parser = new AutoDetectParser();
 
        System.out.println("------------ Parsing an Office Document:");
        extractFromFile(parser, "/demo.doc");
    }
 
    public static void extractFromFile(final Parser parser,
            final String fileName) throws IOException, SAXException,
            TikaException {
        long start = System.currentTimeMillis();
        BodyContentHandler handler = new BodyContentHandler(10000000);
        Metadata metadata = new Metadata();
        InputStream content = AutoDetectionExample.class
                .getResourceAsStream(fileName);
        parser.parse(content, handler, metadata, new ParseContext());
        for (String name : metadata.names()) {
            System.out.println(name + ":\t" + metadata.get(name));
        }
        System.out.println(String.format(
                "------------ Processing took %s millis\n\n",
                System.currentTimeMillis() - start));
        System.out.println("------------ content of document\n" + handler.toString());
    }
}
