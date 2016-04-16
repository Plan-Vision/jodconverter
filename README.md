JODConverter
============

JODConverter (for Java OpenDocument Converter) automates document conversions using LibreOffice or OpenOffice.org.

Forked from https://github.com/nuxeo/jodconverter stating:

VisionR Notes

 - Fixed pdf calc export with very long row count (define print area)
 - Delayed process creation and recreation
 - FIFO process queue replaced with LIFO queue because of the slow full initialization (reuse an old instance before trying to create a new one)
 
