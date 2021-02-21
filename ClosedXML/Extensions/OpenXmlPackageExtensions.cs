// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

namespace DocumentFormat.OpenXml.Packaging
{
    internal static class OpenXmlPackageExtensions
    {
        /// <summary>
        /// Traverse parts in the <see cref="OpenXmlPackage"/> by breadth-first.
        /// </summary>
        public static IEnumerable<OpenXmlPart> GetAllParts(this OpenXmlPackage package)
        {
            if (package is null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            var visited = new HashSet<OpenXmlPart>();
            var queue = new Queue<OpenXmlPart>();

            // Enqueue child parts of the package.
            foreach (var idPartPair in package.Parts)
            {
                queue.Enqueue(idPartPair.OpenXmlPart);
            }

            while (queue.Count > 0)
            {
                var part = queue.Dequeue();

                yield return part;

                foreach (var subIdPartPair in part.Parts)
                {
                    var item = subIdPartPair.OpenXmlPart;

                    if (visited.Add(item))
                    {
                        queue.Enqueue(item);
                    }
                }
            }
        }
    }
}
