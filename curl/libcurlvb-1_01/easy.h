/***************************************************************************
 *
 * Project: libcurl.vb
 *
 * Copyright (c) 2005 Jeff Phillips (jeff@jeffp.net)
 *
 * This software is licensed as described in the file COPYING, which you
 * should have received as part of this distribution.
 *
 * You may opt to use, copy, modify, merge, publish, distribute and/or sell
 * copies of this Software, and permit persons to whom the Software is
 * furnished to do so, under the terms of the COPYING file.
 *
 * This software is distributed on an "AS IS" basis, WITHOUT WARRANTY OF
 * ANY KIND, either express or implied.
 *
 * $Id: easy.h,v 1.1 2005/03/01 00:06:25 jeffreyphillips Exp $
 **************************************************************************/

#pragma once

// obtain inner easy from an exposed easy handle
void* easy_get_inner(void* pvOuter);
// obtain outer easy from inner easy handle
void* easy_get_outer(void* pvInner);
// table creation and removal functions
void easy_create_context_table();
void easy_free_context_table();
