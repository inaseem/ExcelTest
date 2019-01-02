
package ali;

import com.google.gson.annotations.Expose;
import com.google.gson.annotations.SerializedName;

public class PanelizationSummary {

    @SerializedName("containsEpubBubbles")
    @Expose
    private boolean containsEpubBubbles;
    @SerializedName("containsImageBubbles")
    @Expose
    private boolean containsImageBubbles;

    public boolean isContainsEpubBubbles() {
        return containsEpubBubbles;
    }

    public void setContainsEpubBubbles(boolean containsEpubBubbles) {
        this.containsEpubBubbles = containsEpubBubbles;
    }

    public boolean isContainsImageBubbles() {
        return containsImageBubbles;
    }

    public void setContainsImageBubbles(boolean containsImageBubbles) {
        this.containsImageBubbles = containsImageBubbles;
    }

}
