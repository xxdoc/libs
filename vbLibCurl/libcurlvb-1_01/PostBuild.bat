@echo off
rem $Id: PostBuild.bat,v 1.1 2005/03/01 00:06:25 jeffreyphillips Exp $
rem post-build step in Visual Studio IDE to route output files
rem to their proper directories
midl vblibcurl.odl
move vblibcurl.tlb bin

