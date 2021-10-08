/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*
 * This file is part of the LibreOffice project.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * This file incorporates work covered by the following license notice:
 *
 *   Licensed to the Apache Software Foundation (ASF) under one or more
 *   contributor license agreements. See the NOTICE file distributed
 *   with this work for additional information regarding copyright
 *   ownership. The ASF licenses this file to you under the Apache
 *   License, Version 2.0 (the "License"); you may not use this file
 *   except in compliance with the License. You may obtain a copy of
 *   the License at http://www.apache.org/licenses/LICENSE-2.0 .
 */


#include "ListenerHelper.h"
#include "MyProtocolHandler.h"

#include <ctime>
#include <uchar.h>
#include <iostream>
#include <map>
#include <string>

#include <com/sun/star/awt/MessageBoxButtons.hpp>
#include <com/sun/star/awt/Toolkit.hpp>
#include <com/sun/star/awt/XMessageBoxFactory.hpp>
#include <com/sun/star/frame/ControlCommand.hpp>
#include <com/sun/star/text/XTextViewCursorSupplier.hpp>
#include <com/sun/star/sheet/XSpreadsheetView.hpp>
#include <com/sun/star/system/SystemShellExecute.hpp>
#include <com/sun/star/system/SystemShellExecuteFlags.hpp>
#include <com/sun/star/system/XSystemShellExecute.hpp>
#include <cppuhelper/supportsservice.hxx>

#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include <com/sun/star/frame/XComponentLoader.hpp>
#include <com/sun/star/text/XTextDocument.hpp>
#include <com/sun/star/text/XText.hpp>
#include <com/sun/star/text/XTextTable.hpp>
#include <com/sun/star/text/XTextContent.hpp>
#include <com/sun/star/table/XCell.hpp>
#include <com/sun/star/table/XTable.hpp>


using namespace com::sun::star::awt;
using namespace com::sun::star::frame;
using namespace com::sun::star::system;
using namespace com::sun::star::uno;

using namespace com::sun::star::lang;
using namespace com::sun::star::beans;
using namespace com::sun::star::text;
using namespace com::sun::star::table;
using namespace com::sun::star::util;
using rtl::OUString;

using com::sun::star::beans::NamedValue;
using com::sun::star::beans::PropertyValue;
using com::sun::star::sheet::XSpreadsheetView;
using com::sun::star::text::XTextViewCursorSupplier;
using com::sun::star::util::URL;

bool IsCyrillic(OUString word)
{
    int length = word.getLength();
    bool flag_latin = false;
    bool flag_cyrillic = false;
        
    if (length == 0)
        return false;
    
    for (int i = 0; i < length; ++i)
    {
        //std::cout << word[i] << std::endl;
        if (latin.find(word[i]) == std::string::npos)
            if (cyrillic.find(word[i]) == std::string::npos)
                return false;
            else
                flag_cyrillic = true;
        else
            flag_latin = true;
            
    }
    
    return flag_latin && not flag_cyrillic;
}

bool IsLetter(char16_t letter){
    if (bilingual.find(letter) == std::string::npos)
            return false;
    return true;
}

void GenerateText(Reference< XFrame > &mxFrame, int number, int maxlen, int lang)
{
    std::srand(std::time(nullptr));
    
    OUString generated_text = OUString("");
    for (int i = 0; i < number; ++i)
    {
        int cur_len = 1 + rand() % maxlen;
        
        if (i != 0)
        {
            int x_rand = rand() % 4;
            if (x_rand == 1) // 25% chance to gen sep 
            {
                generated_text += OUString(separators[rand() % 6]);
                generated_text += OUString(" ");
            }
            else if (x_rand == 2)
            {
                generated_text += OUString(separators[6]);
            }
            else
            {
                generated_text += OUString(" ");
            }
        }
            
            
        if (lang == 0)
        for (int j = 0; j < cur_len; ++j)
            generated_text += OUString(cyrillic[rand() % 66]);
        
        else if (lang == 1)
        for (int j = 0; j < cur_len; ++j)
            generated_text += OUString(latin[rand() % 52]);
    
    else if (lang == 2)
        for (int j = 0; j < cur_len; ++j)
            generated_text += OUString(bilingual[rand() % 118]);
    }
    int x_rand = rand() % 4;
    if (x_rand == 1 || x_rand == 2) // 50% chance to gen sep 
    {
        generated_text += OUString(separators[rand() % 3]);
        generated_text += OUString(" ");
    }
    else
    {
        generated_text += OUString(" ");
    }
    
    Reference < XComponentLoader > rComponentLoader(mxFrame, UNO_QUERY);    
    Reference <XComponent> xComponent = rComponentLoader->loadComponentFromURL("private:factory/swriter", "_blank", 0, Sequence < PropertyValue > ());
    Reference <XTextDocument> xTextDocument(xComponent, UNO_QUERY);
    Reference < XText > xText = xTextDocument -> getText();
    
    xText->setString(generated_text);
        
}

void HighlightText(Reference< XFrame > &mxFrame)
{
    Reference < XTextDocument > xTextDocument(mxFrame -> getController() -> getModel(), UNO_QUERY);
    Reference < XText > xText = xTextDocument -> getText();
    Reference < XTextCursor > xTextCursor = xText -> createTextCursor();
    int xTextSize = (xText -> getString()).getLength();
    Reference < XPropertySet > xCursorProps(xTextCursor, UNO_QUERY);
    
    int prev_word = 0;
    int cur_point = 0;
    
    while ((prev_word < xTextSize) && (xTextCursor -> goRight(1, true))) 
    {
        cur_point = 0;
        
        while ((cur_point < (xTextSize - prev_word)) && (IsLetter((xTextCursor -> getString())[cur_point]))) 
        {
            xTextCursor -> goRight(1, true);
            ++cur_point;
        }
                
        if (cur_point != (xTextSize - prev_word))
        {
            xTextCursor -> goLeft(1, true);
        }

        if (IsCyrillic(xTextCursor -> getString()))
        {
            xCursorProps -> setPropertyValue("CharColor", makeAny(0xC471ED));
        }
        else
        {
            xCursorProps -> setPropertyValue("CharColor", makeAny(0x000000));
        }
        
        if (xTextCursor -> goRight(1, true)) 
        {
            xTextCursor -> collapseToEnd();
            prev_word += cur_point + 1;
        }
        
        else break;
    }
    xCursorProps -> setPropertyValue("CharColor", makeAny(0x000000));
}

void TableStats(Reference < XFrame > & mxFrame)
{
    Reference < XTextDocument > xTextDocument(mxFrame -> getController() -> getModel(), UNO_QUERY);
    Reference < XText > xText = xTextDocument -> getText();
    Reference < XTextCursor > xTextCursor = xText -> createTextCursor();
    int xTextSize = (xText -> getString()).getLength();
    Reference < XPropertySet > xCursorProps(xTextCursor, UNO_QUERY);
    
    std::map <int, int> len_stats;
    
    int prev_word = 0;
    int cur_point = 0;
    
    while ((prev_word < xTextSize) && (xTextCursor -> goRight(1, true))) 
    {
        cur_point = 0;
        
        while ((cur_point < (xTextSize - prev_word)) && (IsLetter((xTextCursor -> getString())[cur_point]))) 
        {
            xTextCursor -> goRight(1, true);
            ++cur_point;
        }
                
        if (cur_point != (xTextSize - prev_word))
        {
            xTextCursor -> goLeft(1, true);
        }
        
        if (cur_point > 0)
            if (len_stats.find(cur_point) == len_stats.end())
                len_stats[cur_point] = 1;
            else
                ++len_stats[cur_point];
        
        if (xTextCursor -> goRight(1, true)) 
        {
            xTextCursor -> collapseToEnd();
            prev_word += cur_point + 1;
        }
        
        else break;
    }
    
    xTextCursor->gotoEnd(false);
    
// Create a new text table from the document's factory

   Reference<XMultiServiceFactory> oDocMSF (xTextDocument,UNO_QUERY);
   Reference <XTextTable> xTable (oDocMSF->createInstance(
            OUString::createFromAscii("com.sun.star.text.TextTable")),UNO_QUERY);
 
// Specify that we want the table to have N rows and 2 columns

   xTable->initialize(1 + len_stats.size(), 2);
   Reference <XTextRange> xTextRange = xText->getEnd();
 
   Reference <XTextContent> xTextContent (xTable,UNO_QUERY);

// Insert the table into the document
   xText->insertTextContent(xTextRange, xTextContent,(unsigned char) 0);
   Any prop;
   Reference<XCell> xCell = xTable->getCellByName(OUString::createFromAscii("A1"));
   xText = Reference<XText>(xCell,UNO_QUERY);
   xTextCursor = xText->createTextCursor();
   xTextCursor->setString(OUString::createFromAscii("Length of the words"));
   xCell = xTable->getCellByName(OUString::createFromAscii("B1"));   
   xText = Reference<XText>(xCell,UNO_QUERY);
   xTextCursor = xText->createTextCursor();
   Reference<XPropertySet> oCPS(xTextCursor,UNO_QUERY);
   xTextCursor->setString(OUString::createFromAscii("Counts for specific words length"));
   
   
   int i = 2;
   for (auto pair = len_stats.begin(); pair != len_stats.end(); ++pair, ++i)
   {    
    xCell = xTable->getCellByName(OUString::createFromAscii(((char)'A' + std::to_string(i)).c_str()));
    xText = Reference<XText>(xCell,UNO_QUERY);
    xTextCursor = xText->createTextCursor();
    xCell->setValue(pair->first);
    
    xCell = xTable->getCellByName(OUString::createFromAscii(((char)'B' + std::to_string(i)).c_str()));
    xText = Reference<XText>(xCell,UNO_QUERY);
    xTextCursor = xText->createTextCursor();
    Reference<XPropertySet> oCPS(xTextCursor,UNO_QUERY);
    xCell->setValue(pair->second);
   }
      
    
}


ListenerHelper aListenerHelper;

void BaseDispatch::ShowMessageBox( const Reference< XFrame >& rFrame, const ::rtl::OUString& aTitle, const ::rtl::OUString& aMsgText )
{
    if ( !mxToolkit.is() )
        mxToolkit = Toolkit::create(mxContext);
    Reference< XMessageBoxFactory > xMsgBoxFactory( mxToolkit, UNO_QUERY );
    if ( rFrame.is() && xMsgBoxFactory.is() )
    {
        Reference< XMessageBox > xMsgBox = xMsgBoxFactory->createMessageBox(
            Reference< XWindowPeer >( rFrame->getContainerWindow(), UNO_QUERY ),
            com::sun::star::awt::MessageBoxType_INFOBOX,
            MessageBoxButtons::BUTTONS_OK,
            aTitle,
            aMsgText );

        if ( xMsgBox.is() )
            xMsgBox->execute();
    }
}

void BaseDispatch::SendCommand( const com::sun::star::util::URL& aURL, const ::rtl::OUString& rCommand, const Sequence< NamedValue >& rArgs, sal_Bool bEnabled )
{
    Reference < XDispatch > xDispatch =
            aListenerHelper.GetDispatch( mxFrame, aURL.Path );

    FeatureStateEvent aEvent;

    aEvent.FeatureURL = aURL;
    aEvent.Source     = xDispatch;
    aEvent.IsEnabled  = bEnabled;
    aEvent.Requery    = sal_False;

    ControlCommand aCtrlCmd;
    aCtrlCmd.Command   = rCommand;
    aCtrlCmd.Arguments = rArgs;

    aEvent.State <<= aCtrlCmd;
    aListenerHelper.Notify( mxFrame, aEvent.FeatureURL.Path, aEvent );
}

void BaseDispatch::SendCommandTo( const Reference< XStatusListener >& xControl, const URL& aURL, const ::rtl::OUString& rCommand, const Sequence< NamedValue >& rArgs, sal_Bool bEnabled )
{
    FeatureStateEvent aEvent;

    aEvent.FeatureURL = aURL;
    aEvent.Source     = (::com::sun::star::frame::XDispatch*) this;
    aEvent.IsEnabled  = bEnabled;
    aEvent.Requery    = sal_False;

    ControlCommand aCtrlCmd;
    aCtrlCmd.Command   = rCommand;
    aCtrlCmd.Arguments = rArgs;

    aEvent.State <<= aCtrlCmd;
    xControl->statusChanged( aEvent );
}

void SAL_CALL MyProtocolHandler::initialize( const Sequence< Any >& aArguments )
{
    Reference < XFrame > xFrame;
    if ( aArguments.getLength() )
    {
        // the first Argument is always the Frame, as a ProtocolHandler needs to have access
        // to the context in which it is invoked.
        aArguments[0] >>= xFrame;
        mxFrame = xFrame;
    }
}

Reference< XDispatch > SAL_CALL MyProtocolHandler::queryDispatch(   const URL& aURL, const ::rtl::OUString& sTargetFrameName, sal_Int32 nSearchFlags )
{
    Reference < XDispatch > xRet;
    if ( !mxFrame.is() )
        return 0;

    Reference < XController > xCtrl = mxFrame->getController();
    if ( xCtrl.is() && aURL.Protocol == "vnd.demo.complextoolbarcontrols.demoaddon:" )
    {
        Reference < XTextViewCursorSupplier > xCursor( xCtrl, UNO_QUERY );
        Reference < XSpreadsheetView > xView( xCtrl, UNO_QUERY );
        if ( !xCursor.is() && !xView.is() )
            // without an appropriate corresponding document the handler doesn't function
            return xRet;

        if ( 
             aURL.Path == "SpinfieldCmd" ||
             aURL.Path == "EditfieldCmd" ||
             aURL.Path == "DropdownboxCmd" ||
             aURL.Path == "GenWordsBtn" ||
             aURL.Path == "HighlightBtn" ||
             aURL.Path == "TableStatsBtn" )
        {
            xRet = aListenerHelper.GetDispatch( mxFrame, aURL.Path );
            if ( !xRet.is() )
            {
                xRet = xCursor.is() ? (BaseDispatch*) new WriterDispatch( mxContext, mxFrame ) :
                    (BaseDispatch*) new CalcDispatch( mxContext, mxFrame );
                aListenerHelper.AddDispatch( xRet, mxFrame, aURL.Path );
            }
        }
    }

    return xRet;
}

Sequence < Reference< XDispatch > > SAL_CALL MyProtocolHandler::queryDispatches( const Sequence < DispatchDescriptor >& seqDescripts )
{
    sal_Int32 nCount = seqDescripts.getLength();
    Sequence < Reference < XDispatch > > lDispatcher( nCount );

    for( sal_Int32 i=0; i<nCount; ++i )
        lDispatcher[i] = queryDispatch( seqDescripts[i].FeatureURL, seqDescripts[i].FrameName, seqDescripts[i].SearchFlags );

    return lDispatcher;
}

::rtl::OUString MyProtocolHandler_getImplementationName ()
{
    return ::rtl::OUString( MYPROTOCOLHANDLER_IMPLEMENTATIONNAME );
}

Sequence< ::rtl::OUString > SAL_CALL MyProtocolHandler_getSupportedServiceNames(  )
{
    Sequence < ::rtl::OUString > aRet(1);
    aRet[0] = ::rtl::OUString( MYPROTOCOLHANDLER_SERVICENAME );
    return aRet;
}

#undef SERVICE_NAME

Reference< XInterface > SAL_CALL MyProtocolHandler_createInstance( const Reference< XComponentContext > & rSMgr)
{
    return (cppu::OWeakObject*) new MyProtocolHandler( rSMgr );
}

// XServiceInfo
::rtl::OUString SAL_CALL MyProtocolHandler::getImplementationName(  )
{
    return MyProtocolHandler_getImplementationName();
}

sal_Bool SAL_CALL MyProtocolHandler::supportsService( const ::rtl::OUString& rServiceName )
{
    return cppu::supportsService(this, rServiceName);
}

Sequence< ::rtl::OUString > SAL_CALL MyProtocolHandler::getSupportedServiceNames(  )
{
    return MyProtocolHandler_getSupportedServiceNames();
}

void SAL_CALL BaseDispatch::dispatch( const URL& aURL, const Sequence < PropertyValue >& lArgs )
{
    /* It's necessary to hold this object alive, till this method finishes.
       May the outside dispatch cache (implemented by the menu/toolbar!)
       forget this instance during de-/activation of frames (focus!).

        E.g. An open db beamer in combination with the My-Dialog
        can force such strange situation :-(
     */
    Reference< XInterface > xSelfHold(static_cast< XDispatch* >(this), UNO_QUERY);

    if ( aURL.Protocol == "vnd.demo.complextoolbarcontrols.demoaddon:" )
    {
        
        if ( aURL.Path == "HighlightBtn" )
        {
            std::cout << "printed HighlightBtn" << std::endl;
            HighlightText(mxFrame);
        }
        else if ( aURL.Path == "GenWordsBtn" )
        {
            if (words_length.toInt32() <= 0)
                ShowMessageBox(mxFrame, rtl::OUString("Error"), rtl::OUString("Please set all the parameters to generate words."));
            else
                GenerateText(mxFrame, words_number, words_length.toInt32(), language);
        }
        else if ( aURL.Path == "TableStatsBtn" )
        {
            std::cout << "printed TableStatsBtn" << std::endl;
            TableStats(mxFrame);
        }
        else if ( aURL.Path == "SpinfieldCmd" )
        {
            int value;
            for (sal_Int32 i = 0; i < lArgs.getLength(); i++)
                if (lArgs[i].Name == "Value")
                { 
                    lArgs[i].Value >>= value;
                }
            words_number = value;
        }
        else if ( aURL.Path == "DropdownboxCmd" )
        {
            // Retrieve the text argument from the sequence property value
            rtl::OUString aText;
            for ( sal_Int32 i = 0; i < lArgs.getLength(); i++ )
            {
                if ( lArgs[i].Name == "Text" )
                {
                    lArgs[i].Value >>= aText;
                    break;
                }
            }
            if (aText == "Cyrillic")
                language = 0;
            else if (aText == "Latin")
                language = 1;
            else if (aText == "Bilingual")
                language = 2;
        }
    }
}

void SAL_CALL BaseDispatch::addStatusListener( const Reference< XStatusListener >& xControl, const URL& aURL )
{
    if ( aURL.Protocol == "vnd.demo.complextoolbarcontrols.demoaddon:" )
    {
        
        if ( aURL.Path == "SpinfieldCmd" )
        {
            Sequence < NamedValue > aArgs(4);

            // send command to initialize spin button
            aArgs[0].Name = "Value";
            aArgs[0].Value <<= int(0);
            aArgs[1].Name = "LowerLimit";
            aArgs[1].Value <<= int(1);
            aArgs[2].Name = "Step";
            aArgs[2].Value <<= int(1);
            aArgs[3].Name = "OutputFormat";
            aArgs[3].Value <<= rtl::OUString("%d");

            SendCommandTo(xControl, aURL, rtl::OUString("SetValues"), aArgs, sal_True);
        }
        else if ( aURL.Path == "DropdownboxCmd" )
        {
            // A dropdown box is normally used for a group of commands
            // where the user can select one of a defined set.
            Sequence< NamedValue > aArgs( 1 );

            // send command to set context menu content
            
            Sequence< rtl::OUString > aList( 3 );
            aList[0] = "Cyrillic";
            aList[1] = "Latin";
            aList[2] = "Bilingual";

            aArgs[0].Name = "List";
            aArgs[0].Value <<= aList;
            SendCommandTo( xControl, aURL, rtl::OUString( "SetList" ), aArgs, sal_True );
        }
        else if ( aURL.Path == "EditfieldCmd" )
        {
            ::com::sun::star::frame::FeatureStateEvent aEvent;
            aEvent.FeatureURL = aURL;
            aEvent.Source = (::com::sun::star::frame::XDispatch*) this;
            aEvent.IsEnabled = sal_True;
            aEvent.Requery = sal_False;
            aEvent.State = Any();
            xControl->statusChanged( aEvent );
        }

        aListenerHelper.AddListener( mxFrame, xControl, aURL.Path );
    }
}

void SAL_CALL BaseDispatch::removeStatusListener( const Reference< XStatusListener >& xControl, const URL& aURL )
{
    aListenerHelper.RemoveListener( mxFrame, xControl, aURL.Path );
}

void SAL_CALL BaseDispatch::controlEvent( const ControlEvent& Event )
{
    if ( Event.aURL.Protocol == "vnd.demo.complextoolbarcontrols.demoaddon:" )
    {   
        if ( Event.aURL.Path == "SpinfieldCmd" )
        {
            if ( Event.Event == "TextChanged" )
            {
                rtl::OUString aNewText;
                sal_Bool      bHasText( sal_False );
                for (sal_Int32 i = 0; i < Event.aInformation.getLength(); i++)
                    if ( Event.aInformation[i].Name == "Text" )
                    {
                        bHasText = Event.aInformation[i].Value >>= aNewText;
                        break;
                    }
                if ( bHasText )
                {
                    words_number = aNewText.toInt32();
                }
            }
            
        }
        if ( Event.aURL.Path == "EditfieldCmd" )
        {
            // We get notifications whenever the text inside the combobox has been changed.
            // We store the new text into a member.
            if ( Event.Event == "TextChanged" )
            {
                rtl::OUString aNewText;
                sal_Bool      bHasText( sal_False );
                for ( sal_Int32 i = 0; i < Event.aInformation.getLength(); i++ )
                {
                    if ( Event.aInformation[i].Name == "Text" )
                    {
                        bHasText = Event.aInformation[i].Value >>= aNewText;
                        break;
                    }
                }

                if ( bHasText )
                    // words_length = aNewText;
                    words_length = OUString::number(aNewText.toInt32());
            }
        }
    }
}

BaseDispatch::BaseDispatch( const Reference< XComponentContext > &rxContext,
                            const Reference< XFrame >& xFrame,
                            const ::rtl::OUString& rServiceName )
        : mxContext( rxContext )
        , mxFrame( xFrame )
        , msDocService( rServiceName )
        , mbButtonEnabled( sal_True )
{
}

BaseDispatch::~BaseDispatch()
{
    mxFrame.clear();
    mxContext.clear();
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
