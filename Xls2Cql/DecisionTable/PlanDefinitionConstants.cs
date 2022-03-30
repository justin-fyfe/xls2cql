/*
 * Licensed under the Apache License, Version 2.0 (the "License"); you 
 * may not use this file except in compliance with the License. You may 
 * obtain a copy of the License at 
 * 
 * http://www.apache.org/licenses/LICENSE-2.0 
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the 
 * License for the specific language governing permissions and limitations under 
 * the License.
 * 
 * User: Nityan Khanna
 * Date: 2022-3-25
 */

namespace Xls2Cql.DecisionTable
{
    /// <summary>
    /// Represents constants for Plan Definition generation.
    /// </summary>
    internal class PlanDefinitionConstants
    {
        /// <summary>
        /// The action label.
        /// </summary>
        internal const string ActionLabel = "Action";

        /// <summary>
        /// The activity definition canonical url parameter.
        /// </summary>
        internal const string ActivityDefinitionCanonicalUrlParameter = "adCanonicalUrl";

        /// <summary>
        /// The activity definition profile url parameter.
        /// </summary>
        internal const string ActivityDefinitionProfileUrlParameter = "adProfileUrl";

        /// <summary>
        /// The activity definition publisher.
        /// </summary>
        internal const string ActivityDefinitionPublisher = "WHO";

        /// <summary>
        /// The annotations label.
        /// </summary>
        internal const string AnnotationsLabel = "Annotations";

        /// <summary>
        /// The inputs column start.
        /// </summary>
        internal const int InputsColumnStart = 2;

        /// <summary>
        /// The inputs row start.
        /// </summary>
        internal const int InputsRowStart = 8;

        /// <summary>
        /// The inputs label.
        /// </summary>
        internal const string InputsLabel = "Inputs";

        /// <summary>
        /// The plan definition base URL parameter.
        /// </summary>
        internal const string PlanDefinitionBaseUrlParameter = "pdBaseUrl";

        /// <summary>
        /// The output label.
        /// </summary>
        internal const string OutputLabel = "Output";

        /// <summary>
        /// The SNOMED CT description.
        /// </summary>
        internal const string SnomedCtDescription = "Administration of vaccine to produce active immunity";

        /// <summary>
        /// The SNOMED CT URL.
        /// </summary>
        internal const string SnomedCtUrl = "http://snomed.info/sct/";

        /// <summary>
        /// The SNOMED CT vaccination code.
        /// </summary>
        internal const string SnomedCtVaccinationCode = "33879002";
    }
}