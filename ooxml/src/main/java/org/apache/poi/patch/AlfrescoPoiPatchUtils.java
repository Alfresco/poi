/*
 * Copyright (C) 2005-2014 Alfresco Software Limited.
 *
 * This file is part of Alfresco
 *
 * Alfresco is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * Alfresco is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with Alfresco. If not, see <http://www.gnu.org/licenses/>.
 * 
 * 2014 - Alfresco Software, Ltd.
 * Alfresco Software has added this file
 * The details of changes as svn diff can be found in svn at location root/projects/3rd-party/src
 */
package org.apache.poi.patch;

import java.io.Serializable;
import java.util.AbstractList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.POIXMLFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.xslf.usermodel.XSLFFactory;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFootnotes;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFtnEdn;

/**
 * Patch for ALF-17957 and MNT-11823
 * 
 * @author Viachaslau Tsikhanovich
 * @author Dmitry Velichkevich
 */
public class AlfrescoPoiPatchUtils
{
    private static final int DEFAULT_FOOTNOTES_LIMIT = 50;


    private static final String DEFAULT_CONTEXT = "DEFAULT_CONTEXT";

    private static final String PROP_POI_FOOTNOTES_LIMIT = "poiFootnotesLimit";

    private static final String PROP_POI_EXTRACT_PROPERTIES_ONLY = "poiExtractPropertiesOnly";

    private static final String PROP_POI_ALLOWABLE_XSLF_RELATIONSHIP_TYPES = "poiAllowableXslfRelationshipTypes";


    private static final ThreadLocal<String> CONTEXT = new ThreadLocal<String>();


    /**
     * Context specific properties with the configuration of properties for the {@link AlfrescoPoiPatchUtils#DEFAULT_CONTEXT} context
     */
    private static Map<String, Map<String, Serializable>> properties = new HashMap<String, Map<String, Serializable>>();
    static
    {
        Map<String, Serializable> defaultContextProperties = new HashMap<String, Serializable>();
        defaultContextProperties.put(PROP_POI_FOOTNOTES_LIMIT, DEFAULT_FOOTNOTES_LIMIT);
        defaultContextProperties.put(PROP_POI_EXTRACT_PROPERTIES_ONLY, false);
        defaultContextProperties.put(PROP_POI_ALLOWABLE_XSLF_RELATIONSHIP_TYPES, null);
        properties.put(DEFAULT_CONTEXT, defaultContextProperties);
    }


    /**
     * Sets footnotes limit for XWPF documents parser. Default value {@link AlfrescoPoiPatchUtils#DEFAULT_FOOTNOTES_LIMIT} is set if this parameter is not set directly. Related to
     * MNT-577
     * 
     * @param context - {@link String} value which specifies context of the property
     * @param poiFootnotesLimit - {@link Integer} value which determines the desired footnotes limit
     */
    public static void setPoiFootnotesLimit(String context, int poiFootnotesLimit)
    {
        if (poiFootnotesLimit < 0)
        {
            poiFootnotesLimit = DEFAULT_FOOTNOTES_LIMIT;
        }

        setProperty(context, PROP_POI_FOOTNOTES_LIMIT, poiFootnotesLimit);
    }

    /**
     * Sets the flag to determine whether the entire content of XSLF document must be parsed (<code>false</code>) or not (<code>true</code>). Default value is <code>false</code>.
     * Related to MNT-11823
     * 
     * @param context - {@link String} value which specifies context of the property
     * @param poiExtractPropertiesOnly - {@link Boolean} value which determines how parser must process the entire content of XSLF document
     */
    public static void setPoiExtractPropertiesOnly(String context, boolean poiExtractPropertiesOnly)
    {
        setProperty(context, PROP_POI_EXTRACT_PROPERTIES_ONLY, poiExtractPropertiesOnly);
    }

    /**
     * Sets {@link Set}&lt;{@link String}&gt; set which determines the list of allowable relationship types for traversing during analyzing of XSLF document. Default value is
     * <code>null</code>. Usually determines relationships which refer to XSLF configuration parts. Related to MNT-11823
     * 
     * @param context - {@link String} value which specifies context of the property
     * @param poiExtractPropertiesOnly - {@link Set}&lt;{@link String}&gt; instance which contains allowed document relationship types
     */
    public static void setPoiAllowableXslfRelationshipTypes(String context, Set<String> poiAllowableXslfRelationshipTypes)
    {
        setProperty(context, PROP_POI_ALLOWABLE_XSLF_RELATIONSHIP_TYPES, (Serializable) poiAllowableXslfRelationshipTypes);
    }

    /**
     * Thread-safe method to set specified <code>propertyName</code> property. Initializes properties for <code>context</code> if it is not found
     * 
     * @param context - {@link String} value which specifies context of the property
     * @param propertyName - {@link String} value which determines name of the property
     * @param propertyValue - {@link Serializable} instance which specifies value of the property
     */
    private static void setProperty(String context, String propertyName, Serializable propertyValue)
    {
        synchronized (properties)
        {
            Map<String, Serializable> currentProperties = properties.get(context);
            if (null == currentProperties)
            {
                currentProperties = new HashMap<String, Serializable>();
                currentProperties.put(PROP_POI_FOOTNOTES_LIMIT, DEFAULT_FOOTNOTES_LIMIT);
                currentProperties.put(PROP_POI_EXTRACT_PROPERTIES_ONLY, false);
                properties.put(context, currentProperties);
            }
            currentProperties.put(propertyName, propertyValue);
        }
    }

    /**
     * Sets the current POI context. {@link AlfrescoPoiPatchUtils} will use POI properties for specified context
     * 
     * @param context - {@link String} value which specifies current POI context
     */
    public static void setContext(String context)
    {
        CONTEXT.set(context);
    }

    /**
     * Gets current POI context. {@link AlfrescoPoiPatchUtils} will use POI properties for returned context. {@link AlfrescoPoiPatchUtils#DEFAULT_CONTEXT} is used if context is not
     * specified explicitly. Default properties are configured for the default context
     * 
     * @return {@link String} value which determines the current POI context
     */
    private static String getContext()
    {
        String result = CONTEXT.get();
        return (null != result) ? (result) : (DEFAULT_CONTEXT);
    }

    /**
     * Thread-safe method to get <code>propertyName</code> property for the current {@link AlfrescoPoiPatchUtils#getContext()} context
     * 
     * @param <T> - class of the requested property
     * @param valueClass - {@link Class} instance which specifies expected return type of the requested property
     * @param propertyName - {@link String} value which specifies the name of the property
     * @return <code>T</code> instance which represents the value of the property
     */
    @SuppressWarnings("unchecked")
    private static <T> T getProperty(Class<T> valueClass, String propertyName)
    {
        T result = null;
        String context = getContext();

        synchronized (properties)
        {
            Map<String, Serializable> currentProperties = properties.get(context);

            if (null == currentProperties)
            {
                return null;
            }

            Serializable value = currentProperties.get(propertyName);
            if ((null != value) && valueClass.isAssignableFrom(value.getClass()))
            {
                result = (T) value;
            }
        }

        return result;
    }

    /**
     * MNT-577: Alfresco is running 100% CPU for over 10 minutes while extracting metadata for Word office document <br />
     * <br />
     * Converts mutable {@link CTFootnotes} to read-only limited list of {@link CTFtnEdn}
     * 
     * @param mutableFootnotes - {@link CTFootnotes} instance which contains all currently loaded footnotes
     * @return immutable {@link List}&lt;{@link CTFtnEdn}&gt; list
     */
    public static List<CTFtnEdn> getLimitedReadonlyList(final CTFootnotes mutableFootnotes)
    {
        Integer poiFootnotesLimit = getProperty(Integer.class, PROP_POI_FOOTNOTES_LIMIT);

        // on each call parses entire footnotes store to calculate size
        final int originalSize = mutableFootnotes.getFootnoteList().size();
        final int size = Math.min(originalSize, poiFootnotesLimit);

        return new AbstractList<CTFtnEdn>()
        {
            public CTFtnEdn get(int paramInt)
            {
                // on each call parses footnotes store from the beginning to get item
                return mutableFootnotes.getFootnoteArray(paramInt);
            }

            public CTFtnEdn set(int paramInt, CTFtnEdn paramCTFtnEdn)
            {
                throw new IllegalArgumentException("Operation not supported");
            }

            public void add(int paramInt, CTFtnEdn paramCTFtnEdn)
            {
                throw new IllegalArgumentException("Operation not supported");
            }

            public CTFtnEdn remove(int paramInt)
            {
                throw new IllegalArgumentException("Operation not supported");
            }

            public int size()
            {
                return size;
            }
        };
    }

    /**
     * MNT-11823: Upload of PPTX causes very high memory usage leading to system instability<br />
     * <br />
     * Returns filtered list of {@link PackageRelationship} relationships if properties extraction is enabled and parsing is going on for XSLF document or all relationships of the
     * document
     * 
     * @param factory - {@link POIXMLFactory} instance
     * @param packagePart - {@link PackagePart} instance which represents part of OOXML package
     * @return {@link Iterable}&lt;{@link PackageRelationship}&gt; instance which contains list of allowable package relationships
     * @throws InvalidFormatException
     */
    @SuppressWarnings("unchecked")
    public static Iterable<PackageRelationship> getRelationships(POIXMLFactory factory, PackagePart packagePart) throws InvalidFormatException
    {
        Boolean poiExtractPropertiesOnly = getProperty(Boolean.class, PROP_POI_EXTRACT_PROPERTIES_ONLY);
        Set<String> poiAllowableXslfRelationshipTypes = getProperty(Set.class, PROP_POI_ALLOWABLE_XSLF_RELATIONSHIP_TYPES);

        if (!poiExtractPropertiesOnly || !(factory instanceof XSLFFactory))
        {
            return packagePart.getRelationships();
        }

        List<PackageRelationship> result = new LinkedList<PackageRelationship>();

        if (null == poiAllowableXslfRelationshipTypes)
        {
            return result;
        }

        if (1 == poiAllowableXslfRelationshipTypes.size())
        {
            return packagePart.getRelationshipsByType(poiAllowableXslfRelationshipTypes.iterator().next());
        }

        for (String type : poiAllowableXslfRelationshipTypes)
        {
            PackageRelationshipCollection relationshipsByType = packagePart.getRelationshipsByType(type);
            for (PackageRelationship relationship : relationshipsByType)
            {
                result.add(relationship);
            }
        }

        return result;
    }
}
